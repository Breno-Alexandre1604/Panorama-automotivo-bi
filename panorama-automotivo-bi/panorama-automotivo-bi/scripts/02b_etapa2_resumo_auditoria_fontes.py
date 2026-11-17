from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
exec(open('build2.py').read().split("# ================= ajustes em P3")[0].split("wb=load_workbook(SRC)")[0])  # imports/consts
OUT='/mnt/user-data/outputs/Projeto_BI_Automotivo_Etapa2_Eletrificados.xlsx'
wb=load_workbook(OUT)
F='Arial'; NAVY='1B365D'; TEAL='008080'
hdr_fill=PatternFill('solid',fgColor=NAVY); hdr_font=Font(name=F,bold=True,color='FFFFFF',size=10)
teal_fill=PatternFill('solid',fgColor='D9EEEE')
base=Font(name=F,size=10); bold=Font(name=F,size=10,bold=True); inp=Font(name=F,size=10,color='0000FF'); link=Font(name=F,size=10,color='008000')
note=Font(name=F,size=9,italic=True,color='555555')
thin=Side(style='thin',color='BFC5CF'); box=Border(top=thin,bottom=thin,left=thin,right=thin)
UN='#,##0;-#,##0;0'; PC='0.0%'
def title(ws,t,sub=None):
    ws['A1']=t; ws['A1'].font=Font(name=F,size=14,bold=True,color=NAVY)
    if sub: ws['A2']=sub; ws['A2'].font=note
    ws.sheet_view.showGridLines=False
def block(ws,r,t): ws.cell(r,1,t).font=Font(name=F,size=11,bold=True,color=TEAL)
def hdr(ws,r,labels,h=36):
    for i,l in enumerate(labels,1):
        c=ws.cell(r,i,l); c.fill=hdr_fill; c.font=hdr_font; c.border=box; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[r].height=h
def put(ws,r,c,v,font=base,fmt=None,fill=None,wrap=True):
    x=ws.cell(r,c,v); x.font=font; x.border=box; x.alignment=Alignment(wrap_text=wrap,vertical='top')
    if fmt: x.number_format=fmt
    if fill: x.fill=fill
    return x
def widths(ws,w):
    for i,x in enumerate(w,1): ws.column_dimensions[L(i)].width=x

# ---------- LEIA-ME novo
del wb['LEIA-ME']
wl=wb.create_sheet('LEIA-ME',0); wl.sheet_view.showGridLines=False
lines=[('Panorama Automotivo – Etapa 2: Mercado de Eletrificados (2021–2026)',1),
('Disciplina: Análise de Dados e BI – UNIPÊ (Prof. Daniel Brandão). Arquivo consolidado: contém as abas P1–P3 da Etapa 1 (sem alteração de layout) e as novas abas P4–P8.',0),('',0),
('COMO USAR',2),
('RESUMO: respostas rápidas às perguntas 3.2.  Auditoria_Fontes: o que foi conferido, corrigido ou acrescentado na planilha do Luis.  Fontes: endereços e situação de cada fonte.',0),
('Cores: azul = dado digitado de uma fonte; preto = fórmula; verde = vínculo com outra aba; verde-água = destaque ou valor calculado que preenche lacuna.',0),('',0),
('ABAS',2),
('P1_Volume_Anual / P1_Volume_Mensal – emplacamentos de automóveis 2021–2025 (Anuários FENABRAVE).',0),
('P2_Novos_vs_Usados – novos (0 km) × usados estimados.',0),
('P3_Eletrificados – elétricos e híbridos sobre o total de automóveis (anuário FENABRAVE).',0),
('P4_Tecnologia – BEV × híbridos: série 2021–2025 (FENABRAVE) e por tecnologia no critério ABVE (2024, 2025, 1º sem/2026, jun e set/2026); conferência Federação × ABVE.',0),
('P5_Ranking_Modelos – modelos por tecnologia (1º sem/2026), Top 10 geral e Top 10 por família.',0),
('P6_Ranking_Marcas – marcas (soma dos modelos), elétricos por marca (Federação), BYD e Top 15 do anuário FENABRAVE 2023–2025.',0),
('P7_Geografia – regiões, estados e cidades (1º sem/2026, set/2026 e 2025).   P8_Frota_Ref – frota (referência; não é venda).',0),('',0),
('LEIA ANTES DE USAR NOS GRÁFICOS',2),
('1) Dois critérios de "eletrificado". ABVE = BEV + PHEV + HEV + HEV Flex, SEM micro-híbridos (MHEV) desde jan/2025. Anuário/Federação = inclui MHEV. Não misture totais dos dois (ex.: 2025 = 223.912 na ABVE; 282.458 no anuário).',0),
('2) Abrangência: o anuário usado em P1–P3 conta só AUTOMÓVEIS; a ABVE conta veículos LEVES (inclui comerciais leves). Por isso a diferença de 1% a 2% em P4, bloco E.',0),
('3) Período: o escopo do projeto é 2021–2025. Os dados de 2026 (1º semestre, junho e setembro) são complementares e parciais.',0),
('4) Frota ≠ vendas. P8 não deve ser somada nem comparada às demais abas.',0),
('5) Rankings: os do Canal VE/ABVE listam só alguns modelos; o Top 10 e as marcas em P5/P6 são estimativas mínimas (cobertura de BEV 90%, PHEV 82%).',0),
('6) Kaggle: nenhuma base útil de emplacamentos brasileiros por modelo/estado foi encontrada pelo Luis (bases sintéticas ou globais). Não reverifiquei esse ponto.',0),
('7) ABVE Data (painéis) e planilhas da SENATRAN/dados.gov.br não foram abertos diretamente (painéis interativos e domínios .gov.br fora do alcance deste ambiente). Os valores vêm de comunicados oficiais e reportagens que os reproduzem, sempre indicados na coluna Fonte.',0)]
for i,(t,k) in enumerate(lines,1):
    c=wl.cell(i,1,t); c.alignment=Alignment(wrap_text=True,vertical='top')
    c.font=Font(name=F,size=14 if k==1 else 11 if k==2 else 10,bold=bool(k),color=NAVY if k==1 else TEAL if k==2 else '000000')
wl.column_dimensions['A'].width=150

# ---------- RESUMO
wr=wb.create_sheet('RESUMO',1)
title(wr,'Resumo – Aprofundamento no mercado de eletrificados (pergunta 3.2)','Todos os valores abaixo são fórmulas ligadas às abas de dados.')
block(wr,3,'3.2.1 Segmentação por tecnologia')
hdr(wr,4,['Indicador','Valor','Onde'],28)
res1=[('BEV no total de eletrificados – 1º sem/2026 (critério ABVE)','=P4_Tecnologia!G17',PC,'P4, bloco B'),
('PHEV no total – 1º sem/2026','=P4_Tecnologia!H17',PC,'P4, bloco B'),
('HEV + HEV Flex (sem recarga externa) – 1º sem/2026','=P4_Tecnologia!I17+P4_Tecnologia!J17',PC,'P4, bloco B'),
('Plug-in (BEV+PHEV) – 1º sem/2026','=P4_Tecnologia!L17',PC,'P4, bloco B'),
('BEV no total de eletrificados – 2025 (critério ABVE)','=P4_Tecnologia!G16',PC,'P4, bloco B'),
('BEV no total de eletrificados – 2025 (anuário FENABRAVE, automóveis)','=P4_Tecnologia!E9',PC,'P4, bloco A'),
('Eletrificados 1º sem/2026 × 1º sem/2025','=P4_Tecnologia!B27',PC,'P4, bloco C')]
for i,(a,f,fm,o) in enumerate(res1): r=5+i; put(wr,r,1,a); put(wr,r,2,f,link,fm,teal_fill if i==0 else None); put(wr,r,3,o)
block(wr,13,'3.2.2 Ranking de modelos e marcas (1º sem/2026)')
hdr(wr,14,['Indicador','Valor','Onde'],28)
res2=[('Modelo mais vendido','=P5_Ranking_Modelos!B44',None,'P5, bloco A'),('Unidades do 1º colocado','=P5_Ranking_Modelos!E44',UN,'P5, bloco A'),
('2º colocado','=P5_Ranking_Modelos!B45',None,'P5, bloco A'),('3º colocado','=P5_Ranking_Modelos!B46',None,'P5, bloco A'),
('Marca líder (soma dos modelos listados)','=P6_Ranking_Marcas!B5',None,'P6, bloco A'),('Participação da marca líder','=P6_Ranking_Marcas!H5',PC,'P6, bloco A'),
('Participação da BYD nos elétricos jan.–set./2026 (Federação)','=P6_Ranking_Marcas!B'+str(4+15+4+14+2),PC,'P6, bloco C'),
('BYD nos elétricos 2025 (anuário FENABRAVE)','=P6_Ranking_Marcas!E'+str(4+15+4+14+11+2),PC,'P6, bloco D')]
for i,(a,f,fm,o) in enumerate(res2): r=15+i; put(wr,r,1,a); put(wr,r,2,f,link,fm,teal_fill if i==0 else None); put(wr,r,3,o)
block(wr,24,'3.2.3 Distribuição geográfica (emplacamentos)')
hdr(wr,25,['Indicador','Valor','Onde'],28)
res3=[('Sudeste – 1º sem/2026','=P7_Geografia!C5',PC,'P7, bloco A'),('Nordeste – 1º sem/2026','=P7_Geografia!C6',PC,'P7, bloco A'),('Sul – 1º sem/2026','=P7_Geografia!C7',PC,'P7, bloco A'),
('Centro-Oeste – 1º sem/2026','=P7_Geografia!C8',PC,'P7, bloco A'),('Norte – 1º sem/2026','=P7_Geografia!C9',PC,'P7, bloco A'),
('Estado líder (São Paulo) – 1º sem/2026','=P7_Geografia!E15',PC,'P7, bloco B'),('5 maiores estados – 1º sem/2026','=P7_Geografia!E20',PC,'P7, bloco B'),('Cidade líder: São Paulo (un.)','=P7_Geografia!D25',UN,'P7, bloco C')]
for i,(a,f,fm,o) in enumerate(res3): r=26+i; put(wr,r,1,a); put(wr,r,2,f,link,fm,teal_fill if i==0 else None); put(wr,r,3,o)
block(wr,36,'Alertas principais (detalhes em Auditoria_Fontes)')
for i,t in enumerate(['A ABVE e o anuário FENABRAVE contam "eletrificados" de forma diferente (micro-híbridos). Use um critério por gráfico.',
 'O crescimento de 2025 divulgado pela ABVE (+26%) mistura critérios; no mesmo critério (sem MHEV) é +38,9%.',
 'Sul/2025 e Paraná/2025 estavam ausentes na planilha original e foram acrescentados.',
 'Geografia vem de comunicados e reportagens da ABVE; só os 5 maiores estados e cidades foram divulgados (não há as 27 UFs).']):
    c=wr.cell(37+i,1,'• '+t); c.font=base; c.alignment=Alignment(wrap_text=True)
widths(wr,[70,16,22])

# ---------- Auditoria_Fontes
wa=wb.create_sheet('Auditoria_Fontes')
title(wa,'Auditoria das fontes da planilha do Luis','Verificação feita em 10/11/2026 por busca na web e por cruzamento com os Anuários FENABRAVE e com as próprias somas da planilha.')
hdr(wa,4,['Nº','Item','Situação','Evidência','O que foi feito'],30)
G='E2F4E3'; Y='FFF2CC'; R='F8D7DA'; B='DCE9F8'
st={'CONFERIDO':G,'RESSALVA':Y,'ERRO CORRIGIDO':R,'LACUNA PREENCHIDA':B}
items=[
('Totais por tecnologia (junho 47.579; 1º sem 215.023; set 53.126)','CONFERIDO','Somas internas batem; junho, semestre e BEV 90.626 aparecem em Automotive Business, Bem Paraná e Canal VE; set. confere na Gazeta de SP.','Mantido.'),
('Regiões e estados do 1º sem/2026','CONFERIDO','Sudeste 96.159 (44,7%), Nordeste 18,8%, Sul 18%; estados e cidades batem com Bem Paraná. Soma das regiões = 215.022 (1 unidade de arredondamento).','Mantido, nota explicando a unidade.'),
('Total acumulado jan.–set./2026 (381.609)','CONFERIDO','215.023 + jul 56.074 + ago 57.386 + set 53.126 = 381.609.','Conta incluída em P8.'),
('Série "2024 = 177.358 → 2025 = 223.912" (crescimento ≈ +26%)','ERRO CORRIGIDO','A ABVE divulgou 2024 COM micro-híbridos: BEV 61.615 + PHEV 64.009 + HEV 15.271 + HEV Flex 20.277 + MHEV 16.185 = 177.357. O critério sem MHEV só vale desde jan/2025. Logo a comparação mistura critérios.','Mesmo critério (sem MHEV): 161.172 → 223.912 = +38,9%. Corrigido em P4 e P8.'),
('Hipótese "a Federação inclui micro-híbridos" (marcada como não confirmada)','CONFERIDO','A ABVE registra 29.938 MHEV no 1º sem/2026; a diferença nos híbridos (154.472 − 124.397) é 30.075, ou seja, sobram 137 unidades.','Hipótese confirmada em P4, bloco D.'),
('Série de eletrificados do anuário FENABRAVE (P3) × ABVE','CONFERIDO','Anuário 2022: 48.336 × ABVE 49.245; 2023: 92.843 × 93.927; 2024: 174.694 × 177.358; 2025: 282.458 × 285.300 (ABVE com MHEV). Diferença de 1% a 2%: o anuário inclui MHEV.','Descoberta incorporada ao P3 e P4. O 14,1% de P3 é "com micro-híbridos".'),
('Texto do Anuário 2025: "elétricos +224%, híbridos +54%"','RESSALVA','O gráfico do mesmo anuário dá +32% (60.106→79.351) e +77% (114.588→203.107). O Anuário 2024 também cita "5,4%" para eletrificados; o cálculo dá ~9,0%.','Calculado a partir das unidades; texto do anuário não usado.'),
('Fonte FENABRAVE "endereço oficial não localizado"','LACUNA PREENCHIDA','Os Anuários FENABRAVE 2021–2025 (enviados por Breno) trazem elétricos e híbridos por marca e totais anuais. Endereço: fenabrave.org.br/portalv2/Conteudo/anuarios.','Marcas 2023–2025 em P6 (blocos D e E); série anual em P4 (A).'),
('Sul/2025 ausente na tabela de regiões','LACUNA PREENCHIDA','223.912 − (103.964 + 36.596 + 33.964 + 9.303) = 40.085 = 17,9%, igual ao 17,9% divulgado (Vrum).','Calculado em P7 (célula destacada).'),
('Paraná/2025 ausente nos 5 maiores estados','LACUNA PREENCHIDA','Forbes Brasil (ABVE): Paraná 14.024 (6,3%), 5º lugar em 2025.','Acrescentado em P7, bloco B.'),
('Série por tecnologia só existia para 2026','LACUNA PREENCHIDA','ABVE 2025: BEV 80.178, PHEV 101.364, HEV 21.047, HEV Flex 21.323 (soma = 223.912). ABVE 2024 acima.','Acrescentado em P4, bloco B.'),
('Nota do ranking de marcas: "híbridos flex trazem só os 5 primeiros"','ERRO CORRIGIDO','Os 5 modelos listados somam 24.078, exatamente o total de HEV Flex: a lista é completa. BEV cobre 90%, PHEV 82%, HEV 98%.','Cobertura calculada em P5, bloco C.'),
('Ranking por versão do modelo','RESSALVA','O Haval H6 aparece em 4 linhas (PHEV x2, HEV, flex) e Song Plus em 2; por versão ele não entra no Top 10, somado sobe para 2º entre famílias.','Top 10 por família em P5, bloco B.'),
('Percentuais divulgados × calculados (Dolphin Mini 39,47% × 39,4%; BYD set. 64,9% × 65,1%)','RESSALVA','Diferença de denominador na fonte; unidades batem.','Mantido o cálculo sobre o total da planilha; nota nas abas.'),
('Unidades de São Paulo em setembro (26,8% informado)','RESSALVA','Só o percentual foi divulgado; unidades = 26,8% × 53.126 (estimativa). Minas Gerais (2º) sem valor.','Marcado como estimativa em P7.'),
('Kaggle (5 bases descartadas)','RESSALVA','Descrições indicam bases sintéticas ou globais. Não reabri os links nesta rodada.','Sem alteração.'),
('SENATRAN / dados.gov.br (planilhas de frota)','RESSALVA','Páginas oficiais existem, mas as planilhas não foram abertas aqui (domínios .gov.br fora do alcance). Frota vem de NeoCharge e reportagens.','Frota isolada em P8. Para dado oficial por UF, baixar "Frota por UF e combustível" manualmente.'),
('ABVE Data (painéis interativos)','RESSALVA','Os painéis não abrem por leitura automática; números lidos em comunicados da ABVE e reportagens.','Origem indicada linha a linha.'),
('Granularidade geográfica','RESSALVA','Só região, 5 maiores estados e 5 maiores cidades foram divulgados; não existe a distribuição das 27 UFs nas fontes lidas.','Para o mapa completo é preciso o painel da ABVE ou o arquivo da SENATRAN.')]
for i,(it,s,ev,ac) in enumerate(items):
    r=5+i; put(wa,r,1,i+1,fmt='0'); put(wa,r,2,it,bold); put(wa,r,3,s,bold,fill=PatternFill('solid',fgColor=st[s])); put(wa,r,4,ev); put(wa,r,5,ac)
widths(wa,[5,46,20,80,46]); wa.freeze_panes='A5'

# ---------- Fontes (arrumada)
src=load_workbook('/mnt/user-data/uploads/parte_de_Luis_projeto_de_dados.xlsx')['Fontes']
wf=wb.create_sheet('Fontes')
title(wf,'Fontes – nomes, endereços e situação','Consulta original do Luis em 10/11/2026; coluna "Verificação nesta rodada" acrescentada.')
hdr(wf,4,['Nº','Fonte','Tipo','Dado utilizado','Endereço','Situação (consulta do Luis)','Verificação nesta rodada'],32)
ver={'A1':'Não reabri (painel interativo).','A2':'Não reabri (painel interativo).','A3':'Não reabri (painel interativo).','A4':'Números do comunicado reproduzidos por vários veículos: conferido.',
'A5':'Não aberto (.gov.br fora do alcance).','A5a':'Não aberto.','A5b':'Não aberto.','A5c':'Não aberto.','A6':'Não aberto.',
'A7':'SUBSTITUÍDA: usados os Anuários FENABRAVE 2021–2025 (endereço no campo ao lado).','A8':'Não reabri.','B1':'Conferido (ver A4).','B2':'Conferido por busca (setembro/2026).','B15':'Conferido (2025: 223.912; regiões).'}
r=5
for row in src.iter_rows(min_row=5,max_row=44,values_only=True):
    a=row[0]
    if a is None: continue
    if row[1] is None and isinstance(a,str):
        wf.cell(r,1,a).font=Font(name=F,size=11,bold=True,color=TEAL); r+=1; continue
    if a=='A7': row=list(row); row[4]='https://www.fenabrave.org.br/portalv2/Conteudo/anuarios'; row[5]='Site não abriu na consulta do Luis; substituído pelos anuários.'
    vals=list(row[:6])+[ver.get(a,'Não reabri nesta rodada.')]
    for c,v in enumerate(vals,1): put(wf,r,c,v)
    r+=1
r+=1; wf.cell(r,1,'N) Fontes acrescentadas nesta rodada').font=Font(name=F,size=11,bold=True,color=TEAL); r+=1
new=[('N1','Anuários FENABRAVE 2021–2025 (PDF enviados)','Fonte original','Emplacamentos de automóveis; BEV e híbridos por marca e por ano','https://www.fenabrave.org.br/portalv2/Conteudo/anuarios','Lido integralmente (P1–P3, P6)'),
('N2','Canal VE – "Eletrificados crescem 26% em 2025 e somam 224 mil novas vendas" (ABVE)','Reportagem que reproduz a ABVE','2025 por tecnologia; Sul 2025 (por diferença)','https://canalve.com.br/eletrificados-crescem-26-2025-tem-224-mil-novas-vendas-ves/','Lido via busca'),
('N3','Motorshow / Mecânica Online – eletrificados 2024 (ABVE)','Reportagem que reproduz a ABVE','2024 por tecnologia incl. MHEV (177.358)','https://motorshow.com.br/vendas-de-carros-eletricos-e-hibridos-crescem-89-em-2024-no-brasil/','Lido via busca'),
('N4','Autodata – "Vendas de veículos eletrificados crescem 61% em 2025" (ABVE)','Reportagem que reproduz a ABVE','2025 com todas as tecnologias (285,3 mil)','https://www.autodata.com.br/noticias/2026/01/12/vendas-de-veiculos-eletrificados-crescem-61-em-2025/98505/','Lido via busca'),
('N5','Portal N10 – vendas de eletrificados disparam 125% (ABVE)','Reportagem que reproduz a ABVE','MHEV 1º sem/2026 (29.938)','https://portaln10.com.br/automoveis/vendas-de-carros-eletricos-e-hibridos-disparam-125-no-brasil-e-atingem-215-mil-unidades-no-primeiro-semestre-349238/','Lido via busca'),
('N6','Forbes Brasil – carros elétricos e híbridos somam 223 mil vendas em 2025 (ABVE)','Reportagem que reproduz a ABVE','Estados em 2025 (inclui Paraná)','https://forbes.com.br/forbes-life/2026/01/carros-eletricos-e-hibridos-batem-recorde-e-somam-223-mil-vendas-em-2025-no-brasil/','Lido via busca'),
('N7','Vrum – vendas de eletrificados crescem dez vezes mais que a combustão em 2025','Reportagem','Sul 2025 = 17,9%','https://www.vrum.com.br/mercado/2026/01/7328023-vendas-de-eletrificados-crescem-dez-vezes-mais-que-carros-a-combustao-em-2025.html','Lido via busca'),
('N8','Bem Paraná – carros elétricos, vendas 2026 (ABVE)','Reportagem que reproduz a ABVE','Conferência de estados, cidades e tecnologias do 1º sem/2026','https://www.bemparana.com.br/noticias/economia/carros-eletricos-vendas-2026/','Lido via busca'),
('N9','Canal VE – infográfico de eletrificados (ABVE)','Reportagem que reproduz a ABVE','2022 = 49.245; 2023 = 93.927','https://canalve.com.br/infografico-reune-dados-da-atual-situacao-de-eletrificados-no-brasil/','Lido via busca')]
for row in new:
    vals=list(row)+['Acrescentada nesta rodada']
    vals=[vals[0],vals[1],vals[2],vals[3],vals[4],vals[5],'Consulta em 10/11/2026']
    for c,v in enumerate(vals,1): put(wf,r,c,v)
    r+=1
widths(wf,[7,60,24,44,60,50,34]); wf.freeze_panes='A5'

# ---------- ordem das abas
order=['LEIA-ME','RESUMO','P1_Volume_Anual','P1_Volume_Mensal','P2_Novos_vs_Usados','P3_Eletrificados','P4_Tecnologia','P5_Ranking_Modelos','P6_Ranking_Marcas','P7_Geografia','P8_Frota_Ref','Auditoria_Fontes','Fontes']
wb._sheets=[wb[n] for n in order]
# abas de dados da etapa 1: cores de aba
for n in order:
    wb[n].sheet_properties.tabColor = NAVY if n.startswith('P') and n[1] in '123' else TEAL if n.startswith('P') else '888888'
wb.active=1
wb.save(OUT); print('saved')
