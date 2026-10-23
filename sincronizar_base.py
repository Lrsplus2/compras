# -*- coding: utf-8 -*-
"""
Sincronizador da Base CIGAM  —  Solicitação de Compras Nacionais (SCN)

Lê a planilha exportada do CIGAM (Excel .xlsx ou .csv) e envia ao banco da aplicação
SOMENTE as linhas que mudaram (estoque, custo, descrição…) e remove os itens que saíram
da planilha. Pode rodar a cada minuto. Não depende de navegador nem de usuário logado.

Requisitos: Python 3.8+ (somente biblioteca padrão — nada a instalar).

Uso:
    python sincronizar_base.py              # uma execução (use no Agendador de Tarefas, a cada 1 minuto)
    python sincronizar_base.py --loop       # fica rodando e repete a cada 60 s (alternativa ao agendador)
    python sincronizar_base.py --teste      # só lê a planilha e mostra o resumo (não envia nada)
    python sincronizar_base.py --forcar     # reenvia tudo e ignora a proteção de arquivo incompleto
    sincronizar_base.exe --instalar         # (Windows) passa a iniciar sozinho a cada logon e já inicia; sem administrador
    sincronizar_base.exe --desinstalar      # remove a inicialização automática
    Opções: --arquivo=CAMINHO (planilha; padrão Y:\\Planilha CIGAM.xlsx) e --config=PASTA_OU_ARQUIVO.ini (padrão: pasta do programa)

Configuração: arquivo config.ini na mesma pasta (veja config.exemplo.ini).
Mantenha esta pasta (script + config.ini) em local RESTRITO do servidor — o config.ini contém o token.
"""
import sys, os, io, re, json, time, hashlib, zipfile, csv, unicodedata, configparser, urllib.request, urllib.error, datetime
import xml.etree.ElementTree as ET

AQUI = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))  # pasta do .exe / do .py
def _resolver_cfg():
    """config.ini: por padrão na pasta do programa. Pode ser mudado com --config=CAMINHO (arquivo ou pasta)
    ou com a variável de ambiente SCN_CONFIG."""
    alvo = os.environ.get('SCN_CONFIG', '')
    for a in sys.argv[1:]:
        if a.startswith('--config='):
            alvo = a.split('=', 1)[1].strip().strip('"')
    if not alvo:
        return os.path.join(AQUI, 'config.ini')
    return os.path.join(alvo, 'config.ini') if os.path.isdir(alvo) or not alvo.lower().endswith('.ini') else alvo
ARQ_CFG = _resolver_cfg()
ARQ_ESTADO = os.path.join(AQUI, 'estado.json')
ARQ_LOG = os.path.join(AQUI, 'sincronizar_base.log')
ARQ_LOCK = os.path.join(AQUI, 'sincronizar_base.lock')
LOTE = 800          # linhas por envio
LOTE_APAGAR = 2000  # códigos por envio de exclusão


def log(msg):
    linha = '%s  %s' % (datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'), msg)
    try:
        print(linha, flush=True)
    except Exception:
        pass
    try:
        if os.path.exists(ARQ_LOG) and os.path.getsize(ARQ_LOG) > 2_000_000:
            os.replace(ARQ_LOG, ARQ_LOG + '.1')
        with open(ARQ_LOG, 'a', encoding='utf-8') as f:
            f.write(linha + '\n')
    except Exception:
        pass


# ------------------------------------------------------------------ utilidades
def sem_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFD', str(s)) if unicodedata.category(c) != 'Mn')


def nrm(s):  # igual ao nrm() da aplicação: sem acento, minúsculo, sem espaços nas pontas
    return sem_acento('' if s is None else s).lower().strip()


def norm_busca(s):  # igual ao norm() da aplicação: sem acento, MAIÚSCULO
    return sem_acento('' if s is None else s).upper()


def num(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace('R$', '').replace(' ', '')
    if ',' in t:
        t = t.replace('.', '').replace(',', '.')
    try:
        return float(t)
    except ValueError:
        return 0.0


def texto(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return re.sub(r'\s+', ' ', str(v)).strip()


# ------------------------------------------------------------------ leitura da planilha
def col_para_indice(ref):
    letras = re.match(r'[A-Z]+', ref).group(0)
    n = 0
    for c in letras:
        n = n * 26 + (ord(c) - 64)
    return n - 1


def ler_xlsx(caminho):
    """Gera as linhas (listas) da primeira aba. Só biblioteca padrão."""
    ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    with zipfile.ZipFile(caminho) as z:
        nomes = set(z.namelist())
        # primeira aba
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        sheet = wb.find('m:sheets/m:sheet', ns)
        rid = sheet.get('{%s}id' % ns['r'])
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        alvo = None
        for rel in rels:
            if rel.get('Id') == rid:
                alvo = rel.get('Target')
        if alvo is None:
            alvo = 'worksheets/sheet1.xml'
        alvo = alvo.lstrip('/')
        caminho_sheet = alvo if alvo.startswith('xl/') else 'xl/' + alvo
        # textos compartilhados
        comp = []
        if 'xl/sharedStrings.xml' in nomes:
            with z.open('xl/sharedStrings.xml') as f:
                for _, el in ET.iterparse(f, events=('end',)):
                    if el.tag == '{%s}si' % ns['m']:
                        partes = []
                        for t in el.iter('{%s}t' % ns['m']):
                            # ignora fonética (rPh)
                            partes.append(t.text or '')
                        comp.append(''.join(partes))
                        el.clear()
        with z.open(caminho_sheet) as f:
            for _, el in ET.iterparse(f, events=('end',)):
                if el.tag != '{%s}row' % ns['m']:
                    continue
                cels = {}
                for c in el.findall('m:c', ns):
                    ref = c.get('r')
                    if not ref:
                        continue
                    t = c.get('t')
                    v = c.find('m:v', ns)
                    if t == 's':
                        val = comp[int(v.text)] if v is not None and v.text is not None else None
                    elif t == 'inlineStr':
                        val = ''.join((x.text or '') for x in c.iter('{%s}t' % ns['m']))
                    elif t in ('str', 'e'):
                        val = v.text if v is not None else None
                    elif t == 'b':
                        val = bool(int(v.text)) if v is not None and v.text else None
                    else:
                        if v is None or v.text is None:
                            val = None
                        else:
                            try:
                                val = float(v.text)
                            except ValueError:
                                val = v.text
                    cels[col_para_indice(ref)] = val
                el.clear()
                if cels:
                    n = max(cels) + 1
                    yield [cels.get(i) for i in range(n)]
                else:
                    yield []


def ler_csv(caminho):
    with open(caminho, 'rb') as f:
        bruto = f.read()
    for enc in ('utf-8-sig', 'cp1252', 'latin-1'):
        try:
            txt = bruto.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    amostra = txt[:5000]
    delim = ';' if amostra.count(';') >= amostra.count(',') else ','
    for linha in csv.reader(io.StringIO(txt), delimiter=delim):
        yield linha


def ler_planilha(caminho):
    ext = os.path.splitext(caminho)[1].lower()
    if ext in ('.csv', '.txt'):
        return ler_csv(caminho)
    if ext == '.xlsx':
        return ler_xlsx(caminho)
    raise ValueError('Formato não suportado (%s). Salve a planilha como .xlsx ou .csv.' % ext)


# ------------------------------------------------------------------ conversão (mesmas regras da aplicação)
def col_idx(cab, *nomes):
    for n in nomes:
        if n in cab:
            return cab.index(n)
    for n in nomes:
        for i, h in enumerate(cab):
            if h.startswith(n):
                return i
    return -1


def converter(linhas):
    it = iter(linhas)
    cab = None
    for l in it:                     # cabeçalho = primeira linha não vazia
        if any(x not in (None, '') for x in l):
            cab = [nrm(x) for x in l]
            break
    if cab is None:
        raise ValueError('planilha vazia')
    C = dict(
        cod=col_idx(cab, 'codigo do item', 'cod. material', 'cod material', 'codigo'),
        fab=col_idx(cab, 'codigo fabricante', 'cod. fabrica', 'cod fabrica'),
        desc=col_idx(cab, 'descricao', 'descricao/aplicacao'),
        forn=col_idx(cab, 'nome fornecedor', 'fornecedor'),
        un=col_idx(cab, 'unidade', 'un'),
        emp=col_idx(cab, 'empresa'),
        gr=col_idx(cab, 'grupo'),
        est=col_idx(cab, 'estoque previsto', 'estoque atual'),
        sg=col_idx(cab, 'subgrupo'),
        cu=col_idx(cab, 'preco (moeda) compra unit', 'custo unitario', 'preco custo unit'),
        cnpj=col_idx(cab, 'cnpj fornecedor', 'cnpj'))
    if C['cod'] < 0 or C['desc'] < 0:
        raise ValueError('colunas “Cód. Material” e “Descrição” não encontradas no cabeçalho: %s' % ', '.join(cab))

    nomes = {'cod': 'Código do item', 'fab': 'Código fabricante', 'desc': 'Descrição', 'forn': 'Fornecedor', 'un': 'Unidade', 'emp': 'Empresa',
             'gr': 'Grupo', 'est': 'Estoque', 'sg': 'Subgrupo', 'cu': 'Custo unitário', 'cnpj': 'CNPJ Fornecedor'}
    ausentes = [nomes[k] for k in C if C[k] < 0]
    if ausentes:
        log('AVISO: colunas não encontradas na planilha: %s. Cabeçalho lido: %s' % (', '.join(ausentes), ' | '.join(cab)))
    usadas = {nomes[k]: cab[C[k]] for k in C if C[k] >= 0}

    def g(r, k):
        i = C[k]
        return r[i] if 0 <= i < len(r) else None

    itens = {}
    grupos = {}
    forn = {}
    for r in it:
        cod = texto(g(r, 'cod'))
        if not cod:
            continue
        gr = texto(g(r, 'gr'))
        sg = texto(g(r, 'sg'))
        if sg and '·' not in sg and re.match(r'^\d+', gr):
            sg = gr.split(' ')[0] + ' · ' + sg
        fo = texto(g(r, 'forn'))
        cf = texto(g(r, 'fab'))
        d = texto(g(r, 'desc'))
        u = texto(g(r, 'un'))
        e = texto(g(r, 'emp'))
        s = round(num(g(r, 'est')), 4)
        cu = round(num(g(r, 'cu')), 4)
        base = '|'.join([cod, cf, d, fo, u, e, gr, sg, repr(s), repr(cu)])
        itens[cod] = {'codigo': cod, 'codfab': cf, 'd': d, 'fo': fo, 'u': u, 'e': e, 'g': gr, 'sg': sg, 's': s, 'cu': cu,
                      'n': norm_busca(d), 'nf': norm_busca(fo + ' ' + cf), 'h': hashlib.md5(base.encode('utf-8')).hexdigest()[:16]}
        if gr:
            grupos.setdefault(gr, set())
            if sg:
                grupos[gr].add(sg)
        cn = re.sub(r'[^0-9A-Z]', '', texto(g(r, 'cnpj')).upper())
        if len(cn) == 14 and cn != '00000000000000' and fo:
            forn.setdefault(cn, {})
            forn[cn][fo] = forn[cn].get(fo, 0) + 1
    meta = {
        'grupos': {k: sorted(v) for k, v in sorted(grupos.items())},
        'fornecedores': sorted(([c, max(n.items(), key=lambda x: x[1])[0]] for c, n in forn.items()), key=lambda x: x[1]),
        'unidades': sorted({v['u'] for v in itens.values() if v['u']}),
        'fabricantes': sorted({v['fo'] for v in itens.values() if v['fo']}),
        'total': len(itens)}
    meta['_avisos'] = ('Colunas não encontradas: ' + ', '.join(ausentes)) if ausentes else ''
    meta['_usadas'] = usadas
    return itens, meta


# ------------------------------------------------------------------ comunicação com o banco
class Banco:
    def __init__(self, url, chave, token):
        self.url = url.rstrip('/')
        self.chave = chave
        self.token = token

    def rpc(self, nome, corpo, tentativas=4):
        dados = json.dumps(corpo, ensure_ascii=False).encode('utf-8')
        req = urllib.request.Request(self.url + '/rest/v1/rpc/' + nome, data=dados, method='POST', headers={
            'apikey': self.chave, 'Authorization': 'Bearer ' + self.chave, 'Content-Type': 'application/json'})
        ultimo = None
        for t in range(tentativas):
            try:
                with urllib.request.urlopen(req, timeout=120) as r:
                    txt = r.read().decode('utf-8')
                    return json.loads(txt) if txt else None
            except urllib.error.HTTPError as e:
                det = e.read().decode('utf-8', 'replace')[:300]
                if e.code in (401, 403) or 'token inv' in det:
                    raise RuntimeError('Banco recusou o acesso (%s): %s — confira url, chave e token no config.ini' % (e.code, det))
                ultimo = RuntimeError('HTTP %s: %s' % (e.code, det))
            except Exception as e:  # rede
                ultimo = e
            time.sleep(2 * (t + 1))
        raise RuntimeError('falha de comunicação com o banco: %s' % ultimo)


# ------------------------------------------------------------------ execução
PADRAO_PLANILHA = r'Y:\Planilha CIGAM.xlsx'


def _arquivo_planilha(sec):
    """Planilha: --arquivo=CAMINHO > config.ini (arquivo=) > padrão Y:\\Planilha CIGAM.xlsx"""
    for a in sys.argv[1:]:
        if a.startswith('--arquivo='):
            return a.split('=', 1)[1].strip().strip('"')
    v = sec.get('arquivo', '').strip().strip('"')
    return v or PADRAO_PLANILHA


def carregar_cfg():
    if not os.path.exists(ARQ_CFG):  # config.ini deve ficar na mesma pasta do programa
        raise SystemExit('config.ini não encontrado em %s — copie config.exemplo.ini para config.ini e preencha.' % AQUI)
    cp = configparser.ConfigParser(interpolation=None)
    cp.read(ARQ_CFG, encoding='utf-8')
    s = cp['scn']
    return {'url': s['url'].strip(), 'chave': s['chave_publica'].strip(), 'token': s['token'].strip(), 'arquivo': _arquivo_planilha(s), 'arquivo_kit': s.get('arquivo_kit', '').strip().strip('"')}


def ler_estado():
    try:
        with open(ARQ_ESTADO, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def gravar_estado(e):
    tmp = ARQ_ESTADO + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(e, f)
    os.replace(tmp, ARQ_ESTADO)


def agora_iso():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def abrir_copia_estavel(caminho):
    """Copia o arquivo para pasta temporária (evita travar o arquivo do CIGAM e lê só quando estiver estável)."""
    import shutil, tempfile
    for t in range(5):
        try:
            st1 = os.stat(caminho)
            time.sleep(1.0)
            st2 = os.stat(caminho)
            if (st1.st_size, st1.st_mtime) != (st2.st_size, st2.st_mtime):
                continue  # ainda sendo gravado
            tmp = os.path.join(tempfile.gettempdir(), 'scn_base_' + str(os.getpid()) + os.path.splitext(caminho)[1])
            shutil.copyfile(caminho, tmp)
            return tmp, st2
        except (PermissionError, OSError):
            time.sleep(3)
    raise RuntimeError('não foi possível ler o arquivo (em uso ou inacessível): %s' % caminho)


def executar_cigam(forcar=False, teste=False):
    cfg = carregar_cfg()
    banco = Banco(cfg['url'], cfg['chave'], cfg['token'])
    est = ler_estado()
    caminho = cfg['arquivo']
    status = {'verificado_em': agora_iso(), 'ok': True, 'msg': '', 'arquivo': os.path.basename(caminho)}
    try:
        if not os.path.exists(caminho):
            raise RuntimeError('arquivo não encontrado: %s' % caminho)
        st = os.stat(caminho)
        assinatura = [st.st_size, int(st.st_mtime)]
        modificado = datetime.datetime.fromtimestamp(st.st_mtime).strftime('%d/%m/%Y %H:%M:%S')
        if not forcar and not teste and est.get('assinatura') == assinatura:
            status.update(msg='Sem alterações no arquivo', total=est.get('total'), modificado=modificado, alterados=0, ultima_carga=est.get('ultima_carga'))
            banco.rpc('base_meta_set', {'p_token': cfg['token'], 'p_meta': None, 'p_arquivo': os.path.basename(caminho), 'p_status': status})
            log('Sem alterações (%s itens).' % est.get('total'))
            return
        t0 = time.time()
        tmp, st2 = abrir_copia_estavel(caminho)
        try:
            itens, meta = converter(ler_planilha(tmp))
        finally:
            try:
                os.remove(tmp)
            except Exception:
                pass
        avisos = meta.pop('_avisos', '')
        usadas = meta.pop('_usadas', {})
        assinatura = [st2.st_size, int(st2.st_mtime)]
        total = len(itens)
        log('Planilha lida: %d itens, %d fornecedores, %.1f s.' % (total, len(meta['fornecedores']), time.time() - t0))
        if teste:
            amostra = next(iter(itens.values()))
            print(json.dumps(amostra, ensure_ascii=False, indent=1))
            print('Colunas usadas (campo -> cabeçalho da planilha):', json.dumps(usadas, ensure_ascii=False))
            if avisos:
                print('AVISO:', avisos)
            print('grupos:', len(meta['grupos']), '| unidades:', meta['unidades'][:10], '| fabricantes:', len(meta['fabricantes']))
            return
        anterior = est.get('total') or 0
        if not forcar and anterior > 1000 and total < anterior * 0.5:
            raise RuntimeError('o arquivo tem %d itens e a base atual tem %d — parece incompleto; nada foi alterado (use --forcar para ignorar)' % (total, anterior))
        hashes_srv = est.get('hashes')
        if forcar or not hashes_srv:
            hashes_srv = banco.rpc('base_hashes', {'p_token': cfg['token']}) or {}
            log('Hashes do banco carregados: %d itens.' % len(hashes_srv))
        mudou = [v for c, v in itens.items() if hashes_srv.get(c) != v['h']]
        apagar = [c for c in hashes_srv if c not in itens]
        for i in range(0, len(mudou), LOTE):
            banco.rpc('base_sync', {'p_token': cfg['token'], 'p_linhas': mudou[i:i + LOTE]})
            if len(mudou) > LOTE:
                log('  enviados %d/%d' % (min(len(mudou), i + LOTE), len(mudou)))
        for i in range(0, len(apagar), LOTE_APAGAR):
            banco.rpc('base_sync', {'p_token': cfg['token'], 'p_linhas': [], 'p_apagar': apagar[i:i + LOTE_APAGAR]})
        meta_h = hashlib.md5(json.dumps(meta, sort_keys=True, ensure_ascii=False).encode('utf-8')).hexdigest()
        enviar_meta = forcar or est.get('meta_h') != meta_h
        ultima = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        status.update(msg='Sincronizado: %d alterado(s), %d removido(s)' % (len(mudou), len(apagar)) + ((' — ATENÇÃO: ' + avisos) if avisos else ''), total=total, alterados=len(mudou), removidos=len(apagar),
                      modificado=modificado, ultima_carga=ultima)
        banco.rpc('base_meta_set', {'p_token': cfg['token'], 'p_meta': meta if enviar_meta else None,
                                    'p_arquivo': '%s (modificada em %s)' % (os.path.basename(caminho), modificado), 'p_status': status})
        est.update({'assinatura': assinatura, 'total': total, 'hashes': {c: v['h'] for c, v in itens.items()}, 'meta_h': meta_h, 'ultima_carga': ultima})
        gravar_estado(est)
        log('OK: %d alterados, %d removidos, total %d (%.1f s).' % (len(mudou), len(apagar), total, time.time() - t0))
    except Exception as e:
        log('ERRO: %s' % e)
        if not teste:
            try:
                status.update(ok=False, msg=str(e)[:400], total=est.get('total'))
                banco.rpc('base_meta_set', {'p_token': cfg['token'], 'p_meta': None, 'p_arquivo': os.path.basename(caminho), 'p_status': status}, tentativas=1)
            except Exception:
                pass
        raise SystemExit(1)


# ------------------------------------------------------------------ planilha de Kits (Base_Kit.xlsx)
# O banco da aplicação é a fonte dos kits (qualquer usuário cadastra/edita/exclui pela aba "Cadastro de Kit").
# Este programa grava a planilha Base_Kit.xlsx na mesma pasta da Planilha CIGAM sempre que algum kit muda.
# Se alguém editar a planilha diretamente no Excel (e o banco não tiver alterações pendentes), as mudanças são lidas de volta.
CAB_KIT = ['Grupo', 'Kit', 'Descrição do kit', 'Código do kit', 'Modelo', 'Código CIGAM', 'Descrição do item (CIGAM)', 'Qtde', 'Fornecimento', 'Equipamento']
_ILEGAL = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f]')


def _esc_xml(v):
    return _ILEGAL.sub('', str(v)).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def _letra(i):
    r = ''
    i += 1
    while i:
        i, m = divmod(i - 1, 26)
        r = chr(65 + m) + r
    return r


def escrever_xlsx(caminho, cab, linhas, nome_aba='Kits'):
    larg = [18, 16, 46, 15, 22, 15, 60, 9, 16, 14]
    def cel(ref, v, est=''):
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            f = float(v)
            return '<c r="%s"%s><v>%s</v></c>' % (ref, est, str(int(f)) if f.is_integer() else repr(f))
        return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, est, _esc_xml('' if v is None else v))
    linhas_xml = ['<row r="1">' + ''.join(cel(_letra(i) + '1', c, ' s="1"') for i, c in enumerate(cab)) + '</row>']
    for n, lin in enumerate(linhas, start=2):
        linhas_xml.append('<row r="%d">' % n + ''.join(cel(_letra(i) + str(n), v) for i, v in enumerate(lin)) + '</row>')
    ult = _letra(len(cab) - 1) + str(len(linhas) + 1)
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
             '<cols>' + ''.join('<col min="%d" max="%d" width="%d" customWidth="1"/>' % (i + 1, i + 1, w) for i, w in enumerate(larg[:len(cab)])) + '</cols>'
             '<sheetData>' + ''.join(linhas_xml) + '</sheetData><autoFilter ref="A1:%s"/></worksheet>' % ult)
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="%s" sheetId="1" r:id="rId1"/></sheets></workbook>' % _esc_xml(nome_aba))
    wbrels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
              '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
              '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>'
              '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
              '<fill><patternFill patternType="solid"><fgColor rgb="FF0D2A4A"/><bgColor indexed="64"/></patternFill></fill></fills>'
              '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
              '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>'
              '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    tmp = caminho + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels); z.writestr('xl/workbook.xml', wb)
        z.writestr('xl/_rels/workbook.xml.rels', wbrels); z.writestr('xl/worksheets/sheet1.xml', sheet); z.writestr('xl/styles.xml', styles)
    try:
        os.replace(tmp, caminho)
    except OSError as e:
        try:
            os.remove(tmp)
        except Exception:
            pass
        raise RuntimeError('não foi possível gravar %s (aberta no Excel por alguém?): %s' % (caminho, e))


def _t(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def ler_kits_xlsx(caminho):
    """Lê a planilha no formato gravado por este programa (uma linha por item). Devolve lista de kits."""
    tmp, _ = abrir_copia_estavel(caminho)
    try:
        linhas = list(ler_xlsx(tmp))
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass
    cab = None
    for i, l in enumerate(linhas):
        nomes = [nrm(_t(c)) for c in l]
        if 'grupo' in nomes and 'kit' in nomes:
            cab = nomes; ini = i + 1; break
    if cab is None:
        raise ValueError('formato não reconhecido: a primeira linha precisa ter as colunas Grupo e Kit')
    def ix(*ns):
        for n in ns:
            if n in cab:
                return cab.index(n)
        return None
    ig, ik, idk, ick, im, ic, iq, ifo, ie = ix('grupo'), ix('kit'), ix('descricao do kit'), ix('codigo do kit'), ix('modelo'), ix('codigo cigam'), ix('qtde', 'qtd', 'quantidade'), ix('fornecimento'), ix('equipamento')
    kits, ordem = {}, []
    for l in linhas[ini:]:
        g = _t(l[ig]) if ig is not None and ig < len(l) else ''
        k = _t(l[ik]) if ik is not None and ik < len(l) else ''
        if not g or not k:
            continue
        def c(i):
            return l[i] if i is not None and i < len(l) else None
        chave = (g, k)
        if chave not in kits:
            kits[chave] = {'grupo': g, 'nome': k, 'descricao': _t(c(idk)), 'codigo_kit': _t(c(ick)), 'itens': []}
            ordem.append(chave)
        cod = c(ic)
        if isinstance(cod, float) and cod.is_integer() and len(str(int(cod))) < 12:
            cod = str(int(cod)).zfill(12)
        cod, mod = _t(cod), _t(c(im))
        if cod or mod:
            q = c(iq)
            try:
                q = float(str(q).replace(',', '.')) if q not in (None, '') else 0
            except ValueError:
                q = 0
            kits[chave]['itens'].append({'codigo': cod, 'modelo': mod, 'qtd': q, 'fornecimento': _t(c(ifo)), 'equipamento': _t(c(ie))})
    return [kits[c] for c in ordem]


def _md5_arq(caminho):
    h = hashlib.md5()
    with open(caminho, 'rb') as f:
        for bloco in iter(lambda: f.read(1 << 20), b''):
            h.update(bloco)
    return h.hexdigest()


def _exportar_kits(banco, tok, arq):
    r = banco.rpc('kit_sync_exportar', {'p_token': tok}) or {}
    linhas = r.get('linhas') or []
    escrever_xlsx(arq, CAB_KIT, linhas)
    h = _md5_arq(arq)
    nk = len({(l[0], l[1]) for l in linhas})
    banco.rpc('kit_sync_marcar', {'p_token': tok, 'p_versao': int(r.get('versao') or 0), 'p_arquivo': arq, 'p_hash': h, 'p_erro': ''})
    log('Kits: planilha gravada (%d kit(s), %d linha(s)) em %s' % (nk, len(linhas), arq))


def ciclo_kits(cfg, banco, forcar=False):
    tok = cfg['token']
    arq = cfg.get('arquivo_kit') or os.path.join(os.path.dirname(cfg['arquivo']), 'Base_Kit.xlsx')
    try:
        est = banco.rpc('kit_sync_estado', {'p_token': tok}) or {}
        versao = int(est.get('versao') or 0)
        exp = est.get('exportada_versao')
        nk_banco = int(est.get('n_kits') or 0)
        sujo = forcar or exp is None or int(exp) != versao
        existe = os.path.exists(arq)
        h_arq = None
        if existe:
            try:
                h_arq = _md5_arq(arq)
            except OSError as e:
                log('Kits: planilha em uso, tentarei no próximo ciclo (%s)' % e)
                return
        alterada_fora = existe and h_arq != (est.get('exportada_hash') or '')
        # 1) alguém editou o Excel e o banco não mudou -> importa do Excel
        # 2) primeira vez, banco sem kits e planilha existente -> importa do Excel
        importar = existe and alterada_fora and ((not sujo) or (exp is None and nk_banco == 0 and not forcar))
        if importar:
            try:
                kits = ler_kits_xlsx(arq)
            except Exception as e:
                msg = 'Planilha de Kits não pôde ser lida (%s). Nada foi alterado.' % e
                log('Kits: ' + msg)
                banco.rpc('kit_sync_marcar', {'p_token': tok, 'p_versao': None, 'p_arquivo': arq, 'p_hash': None, 'p_erro': msg[:400]})
                return
            if nk_banco >= 5 and len(kits) < nk_banco * 0.5:
                msg = 'A planilha de Kits tem %d kit(s) e o banco tem %d — parece incompleta; nada foi alterado (use --forcar para gravar o banco na planilha).' % (len(kits), nk_banco)
                log('Kits: ' + msg)
                banco.rpc('kit_sync_marcar', {'p_token': tok, 'p_versao': None, 'p_arquivo': arq, 'p_hash': None, 'p_erro': msg[:400]})
                return
            n = banco.rpc('kit_sync_importar', {'p_token': tok, 'p': kits})
            log('Kits: %s kit(s) importado(s) da planilha editada no Excel.' % n)
            _exportar_kits(banco, tok, arq)
            return
        if sujo or not existe:
            if existe and alterada_fora:  # os dois mudaram: o banco vence, a planilha vira cópia de segurança
                import shutil
                bkp = os.path.splitext(arq)[0] + '_copia_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '.xlsx'
                shutil.copyfile(arq, bkp)
                log('Kits: planilha e banco mudaram ao mesmo tempo — o banco prevaleceu; cópia da planilha em %s' % bkp)
            _exportar_kits(banco, tok, arq)
            return
        banco.rpc('kit_sync_marcar', {'p_token': tok, 'p_versao': None, 'p_arquivo': arq, 'p_hash': None, 'p_erro': ''})
    except Exception as e:
        log('Kits ERRO: %s' % e)
        try:
            banco.rpc('kit_sync_marcar', {'p_token': tok, 'p_versao': None, 'p_arquivo': arq, 'p_hash': None, 'p_erro': str(e)[:400]}, tentativas=1)
        except Exception:
            pass


def executar(forcar=False, teste=False):
    erro = None
    try:
        executar_cigam(forcar=forcar, teste=teste)
    except SystemExit as e:
        erro = e
    if not teste:
        try:
            cfg = carregar_cfg()
            ciclo_kits(cfg, Banco(cfg['url'], cfg['chave'], cfg['token']), forcar=forcar)
        except SystemExit:
            pass
        except Exception as e:
            log('Kits ERRO: %s' % e)
    if erro:
        raise erro



def com_trava(fn):
    """Evita duas execuções ao mesmo tempo (agendador a cada minuto)."""
    try:
        if os.path.exists(ARQ_LOCK) and time.time() - os.path.getmtime(ARQ_LOCK) < 600:
            log('Execução anterior ainda em andamento — pulando.')
            return
        with open(ARQ_LOCK, 'w') as f:
            f.write(str(os.getpid()))
        fn()
    finally:
        try:
            os.remove(ARQ_LOCK)
        except Exception:
            pass


CHAVE_RUN = r'Software\Microsoft\Windows\CurrentVersion\Run'
NOME_RUN = 'SCN_SincronizarBase'
ARQ_PID = os.path.join(AQUI, 'sincronizar_base.loop')


def _comando_base():
    if getattr(sys, 'frozen', False):
        return '"%s"' % sys.executable
    return '"%s" "%s"' % (sys.executable, os.path.abspath(__file__))


def _extras():
    ex = ''
    for a in sys.argv[1:]:
        if a.startswith('--config=') or a.startswith('--arquivo='):
            ex += ' "%s"' % a
    return ex


def _admin():
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


NOME_TAREFA = 'SCN_SincronizarBase'


def _em_rede(caminho):
    """True se o caminho está em unidade de rede/mapeada (a conta SYSTEM não consegue usá-la)."""
    try:
        if caminho.startswith('\\\\'):
            return True
        import ctypes
        return ctypes.windll.kernel32.GetDriveTypeW(os.path.splitdrive(caminho)[0] + '\\') == 4
    except Exception:
        return False


def instalar():
    """Com administrador: cria tarefa agendada que inicia com o Windows, como SYSTEM, SEM limite de tempo e com reinício automático.
    Sem administrador: registra a inicialização no logon do usuário atual."""
    import subprocess
    exe = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
    pasta = os.path.dirname(exe)
    args = '--loop' + _extras()
    if _admin():
        if _em_rede(exe):
            print('ATENÇÃO: o programa está em unidade de rede/mapeada (%s). A conta SYSTEM não consegue iniciá-lo de lá. '
                  'Copie sincronizar_base.exe e config.ini para uma pasta LOCAL do servidor (ex.: C:\\SCN) e rode --instalar de lá.' % pasta)
        q = lambda x: x.replace("'", "''")
        ps = ("$a = New-ScheduledTaskAction -Execute '%s' -Argument '%s' -WorkingDirectory '%s';"
              "$t = New-ScheduledTaskTrigger -AtStartup;"
              "$s = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) -StartWhenAvailable -MultipleInstances IgnoreNew;"
              "Register-ScheduledTask -TaskName '%s' -Action $a -Trigger $t -Settings $s -User 'SYSTEM' -RunLevel Highest -Force | Out-Null;"
              "Start-ScheduledTask -TaskName '%s'") % (q(exe), q(args.replace('"', '\\"')), q(pasta), NOME_TAREFA, NOME_TAREFA)
        r = subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', ps], capture_output=True, text=True)
        if r.returncode == 0:
            print('Instalado como tarefa do Windows (usuário SYSTEM), sem limite de tempo e com reinício automático.\n'
                  'Inicia com o computador e roda sem precisar de login. A planilha deve estar acessível pela conta SYSTEM: use caminho local ou caminho de rede completo '
                  '(\\\\servidor\\pasta\\arquivo.xlsx) em "arquivo =" no config.ini, nunca uma letra de unidade mapeada.\n'
                  'Acompanhe pelo sincronizar_base.log e pela tela "Base CIGAM".')
            return
        print('Não foi possível criar a tarefa: ' + (r.stderr or r.stdout).strip())
    import winreg
    cmd2 = _comando_base() + ' --loop --oculto' + _extras()
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, CHAVE_RUN) as k:
        winreg.SetValueEx(k, NOME_RUN, 0, winreg.REG_SZ, cmd2)
    print('Instalado SEM administrador: inicia sozinho sempre que ESTE usuário entrar no Windows e só roda enquanto a sessão estiver aberta. '
          'Para rodar independentemente de usuário, execute --instalar como administrador.\nComando: ' + cmd2)
    subprocess.Popen(cmd2, shell=False, creationflags=0x08000000 | 0x00000008, close_fds=True)
    print('Sincronização iniciada em segundo plano.')


def desinstalar():
    import subprocess
    r = subprocess.run(['schtasks', '/delete', '/tn', NOME_TAREFA, '/f'], capture_output=True, text=True)  # funciona para tarefas criadas por qualquer método
    if r.returncode == 0:
        print('Tarefa do Windows removida.')
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, CHAVE_RUN, 0, winreg.KEY_SET_VALUE) as k:
            winreg.DeleteValue(k, NOME_RUN)
        print('Inicialização no logon removida.')
    except Exception:
        pass
    print('Se o programa estiver rodando, encerre o processo sincronizar_base.exe no Gerenciador de Tarefas.')


def _outra_instancia_ativa():
    try:
        return time.time() - os.path.getmtime(ARQ_PID) < 600
    except OSError:
        return False


def _batimento():
    try:
        with open(ARQ_PID, 'w') as f:
            f.write(str(os.getpid()))
    except Exception:
        pass


if __name__ == '__main__':
    args = set(sys.argv[1:])
    forcar, teste, loop = '--forcar' in args, '--teste' in args, '--loop' in args
    if '--instalar' in args:
        instalar()
    elif '--desinstalar' in args:
        desinstalar()
    elif loop:
        if '--oculto' in args and '--filho' not in args:  # reabre sem janela e encerra este
            import subprocess
            subprocess.Popen(_comando_base() + ' --loop --filho' + _extras(), creationflags=0x08000000 | 0x00000008, close_fds=True)
            sys.exit(0)
        if _outra_instancia_ativa():
            log('Já existe uma instância em execução — encerrando esta.')
            sys.exit(0)
        log('Modo contínuo: verificando a cada 60 s. Ctrl+C para sair.')
        while True:
            _batimento()
            try:
                com_trava(lambda: executar(forcar=forcar))
            except SystemExit:
                pass
            except Exception as e:
                log('ERRO inesperado: %s' % e)
            forcar = False
            time.sleep(60)
    elif teste:
        executar(teste=True)
    else:
        com_trava(lambda: executar(forcar=forcar))
