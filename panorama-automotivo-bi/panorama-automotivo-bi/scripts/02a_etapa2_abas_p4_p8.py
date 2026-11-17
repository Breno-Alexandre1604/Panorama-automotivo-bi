from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from data import *

SRC='/mnt/user-data/outputs/FENABRAVE_Automoveis_2021-2025_Etapa1.xlsx'
OUT='/mnt/user-data/outputs/Projeto_BI_Automotivo_Etapa2_Eletrificados.xlsx'
wb=load_workbook(SRC)

F='Arial'
NAVY='1B365D'; TEAL='008080'
hdr_fill=PatternFill('solid',fgColor=NAVY); hdr_font=Font(name=F,bold=True,color='FFFFFF',size=10)
teal_fill=PatternFill('solid',fgColor='D9EEEE'); grey=PatternFill('solid',fgColor='F4F6F9')
base=Font(name=F,size=10); inp=Font(name=F,size=10,color='0000FF'); link=Font(name=F,size=10,color='008000'); bold=Font(name=F,size=10,bold=True)
note=Font(name=F,size=9,italic=True,color='555555')
thin=Side(style='thin',color='BFC5CF'); box=Border(top=thin,bottom=thin,left=thin,right=thin)
UN='#,##0;-#,##0;0'; PC='0.0%'

def title(ws,t,sub=None):
    ws['A1']=t; ws['A1'].font=Font(name=F,size=14,bold=True,color=NAVY)
    if sub: ws['A2']=sub; ws['A2'].font=note
    ws.sheet_view.showGridLines=False
def block(ws,r,t):
    c=ws.cell(r,1,t); c.font=Font(name=F,size=11,bold=True,color=TEAL)
def hdr(ws,r,labels,h=44):
    for i,l in enumerate(labels,1):
        c=ws.cell(r,i,l); c.fill=hdr_fill; c.font=hdr_font; c.border=box
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[r].height=h
def put(ws,r,c,v,font=base,fmt=None,fill=None,wrap=False,align=None):
    x=ws.cell(r,c,v); x.font=font; x.border=box
    if fmt: x.number_format=fmt
    if fill: x.fill=fill
    if wrap or align: x.alignment=Alignment(wrap_text=wrap,vertical='center',horizontal=align)
    return x
def txt(ws,r,t,font=note,col=1):
    c=ws.cell(r,col,t); c.font=font
def widths(ws,w):
    for i,x in enumerate(w,1): ws.column_dimensions[L(i)].width=x

# ================= ajustes em P3 (compatibilidade com a nova evidência) =================
p3=wb['P3_Eletrificados']
p3['A8']='Nota: % sobre o total de emplacamentos de automóveis (SENATRAN). "Híbridos" do anuário FENABRAVE inclui micro-híbridos (MHEV) – ver P4, bloco E. Os gráficos de combustível do anuário usam outra base (13,2% em 2025; 10,30% em 2024) e o Anuário 2024 cita 5,4%.'
p3['A10']='Atenção: o texto do Anuário 2025 diz "elétricos +224% e híbridos +54%", mas o próprio gráfico dá +32% (60.106→79.351) e +77% (114.588→203.107). Os números desta aba seguem o gráfico.'
p3['A10'].font=Font(name=F,size=9,italic=True)

# ================= P4_Tecnologia =================
ws=wb.create_sheet('P4_Tecnologia')
title(ws,'P4 – Segmentação por tecnologia: BEV × híbridos (PHEV, HEV)','Pergunta 3.2.1. Duas visões: (A) série anual 2021–2025 do anuário FENABRAVE (automóveis) e (B) critério ABVE por tecnologia (veículos leves).')
block(ws,3,'A) Série anual 2021–2025 – Anuário FENABRAVE (somente automóveis; "híbridos" = PHEV + HEV + micro-híbridos)')
hdr(ws,4,['Ano','Elétricos BEV (un.)','Híbridos (un.)','Total eletrificados (un.)','% BEV','% Híbridos','Cresc. BEV a.a.','Cresc. Híbridos a.a.','Fonte'])
for i in range(5):
    r=5+i; s=2+i
    put(ws,r,1,2021+i,fmt='0')
    put(ws,r,2,f"=P3_Eletrificados!B{s}",link,UN); put(ws,r,3,f"=P3_Eletrificados!C{s}",link,UN)
    put(ws,r,4,f"=B{r}+C{r}",fmt=UN); put(ws,r,5,f"=B{r}/D{r}",fmt=PC,fill=teal_fill); put(ws,r,6,f"=C{r}/D{r}",fmt=PC)
    put(ws,r,7,'' if i==0 else f"=B{r}/B{r-1}-1",fmt=PC); put(ws,r,8,'' if i==0 else f"=C{r}/C{r-1}-1",fmt=PC)
    put(ws,r,9,'Aba P3_Eletrificados (Anuários 2023 e 2025)')
txt(ws,10,'Em 2025 os híbridos (inclui micro-híbridos) já respondem por cerca de 72% dos eletrificados do anuário; veja a conferência com a ABVE no bloco E.')

block(ws,13,'B) Critério ABVE por tecnologia (veículos leves, inclui comerciais leves) – a ABVE exclui os micro-híbridos (MHEV) desde jan/2025')
hdr(ws,14,['Período','BEV','PHEV','HEV (gasolina)','HEV Flex (etanol)','Eletrificados ABVE (sem MHEV)','% BEV','% PHEV','% HEV','% HEV Flex','Plug-in (BEV+PHEV)','% Plug-in','Total divulgado pela ABVE','Diferença (soma − divulgado)','MHEV (informativo)','Fonte'],56)
rowsB=[('2024 (critério antigo, divulgado com MHEV)',61615,64009,15271,20277,177358,16185,'ABVE via Motorshow/Mecânica Online (jan/2025)'),
('2025 (ano)',80178,101364,21047,21323,223912,'=C47-M16','ABVE via Canal VE (06/01/2026); MHEV por diferença com Autodata (aprox.)'),
('1º semestre/2026',90626,76400,23919,24078,215023,29938,'ABVE – comunicado de 06/07/2026 (aba do Luis); MHEV: Portal N10'),
('Junho/2026',21138,18206,4507,3728,47579,None,'ABVE – comunicado de 06/07/2026'),
('Setembro/2026',22246,20202,4892,5786,53126,None,'ABVE via Gazeta de São Paulo (out/2026)')]
for i,(per,b,p,h,f,tot,mh,src) in enumerate(rowsB):
    r=15+i
    put(ws,r,1,per,bold)
    for c,v in zip((2,3,4,5),(b,p,h,f)): put(ws,r,c,v,inp,UN)
    put(ws,r,6,f"=SUM(B{r}:E{r})",fmt=UN)
    for c,col in zip((7,8,9,10),'BCDE'): put(ws,r,c,f"={col}{r}/F{r}",fmt=PC,fill=teal_fill if c==7 else None)
    put(ws,r,11,f"=B{r}+C{r}",fmt=UN); put(ws,r,12,f"=K{r}/F{r}",fmt=PC)
    put(ws,r,13,tot,inp,UN); put(ws,r,14,f"=F{r}-M{r}",fmt=UN)
    put(ws,r,15,mh if mh is not None else '',(base if isinstance(mh,str) else inp),UN)
    put(ws,r,16,src,wrap=True)
    ws.row_dimensions[r].height=30
txt(ws,20,'Em 2024 a diferença (−16.186) é o MHEV: o total 177.358 divulgado inclui 16.185 micro-híbridos. Por isso "177.358 → 223.912 (+26%)" mistura critérios; sem MHEV a base de 2024 é 161.172 (ver P8, bloco E).')
txt(ws,21,'HEV = híbrido convencional a gasolina; HEV Flex = híbrido a etanol/gasolina. Ambos sem recarga externa. PHEV tem recarga externa.')

block(ws,23,'C) Crescimento e participação – 1º semestre/2026 × 1º semestre/2025 (ABVE)')
hdr(ws,24,['Indicador','Valor','','','Fonte'],30)
c_rows=[('Eletrificados 1º sem/2025 (un.)',95493,UN,inp,'Canal VE – 1º semestre 2026 (ABVE)'),
('Eletrificados 1º sem/2026 (un.)','=F17',UN,base,'Bloco B'),
('Crescimento 1º sem/2026 sobre 1º sem/2025','=B26/B25-1',PC,base,'Calculado (a ABVE divulgou +125%)'),
('Participação nas vendas de leves – 1º sem/2025',0.097,PC,inp,'ABVE via Canal VE'),
('Participação nas vendas de leves – 1º sem/2026',0.158,PC,inp,'ABVE via Canal VE (arredondada para 16% no comunicado)')]
for i,(a,v,fm,fo,s) in enumerate(c_rows):
    r=25+i; put(ws,r,1,a); put(ws,r,2,v,fo,fm); put(ws,r,3,''); put(ws,r,4,''); put(ws,r,5,s)

block(ws,32,'D) Conferência 1º sem/2026 – Federação (FENABRAVE) × ABVE: a diferença nos híbridos é o micro-híbrido')
hdr(ws,33,['Indicador','Federação (FENABRAVE)','ABVE','Diferença','Fonte'],36)
put(ws,34,1,'Elétricos 100% a bateria'); put(ws,34,2,90470,inp,UN); put(ws,34,3,'=B17',link,UN); put(ws,34,4,'=B34-C34',fmt=UN); put(ws,34,5,'BP Money (Federação) × ABVE')
put(ws,35,1,'Híbridos (todos os tipos)'); put(ws,35,2,154472,inp,UN); put(ws,35,3,'=C17+D17+E17',link,UN); put(ws,35,4,'=B35-C35',fmt=UN); put(ws,35,5,'BP Money (Federação) × ABVE')
put(ws,36,1,'TOTAL',bold); put(ws,36,2,'=B34+B35',fmt=UN); put(ws,36,3,'=C34+C35',fmt=UN); put(ws,36,4,'=B36-C36',fmt=UN); put(ws,36,5,'Correio Braziliense: "244 mil" (Federação)')
put(ws,37,1,'% de elétricos a bateria no total'); put(ws,37,2,'=B34/B36',fmt=PC); put(ws,37,3,'=C34/C36',fmt=PC); put(ws,37,4,''); put(ws,37,5,'Calculado')
put(ws,38,1,'MHEV – micro-híbridos (ABVE, informativo)'); put(ws,38,2,''); put(ws,38,3,'=O17',link,UN); put(ws,38,4,''); put(ws,38,5,'Portal N10 (ABVE): 29.938 no semestre')
put(ws,39,1,'Diferença dos híbridos menos MHEV (sobra)',bold); put(ws,39,2,''); put(ws,39,3,''); put(ws,39,4,'=D35-C38',bold,UN,teal_fill); put(ws,39,5,'Sobra de ~137 unidades: a hipótese do micro-híbrido se confirma')

block(ws,42,'E) Anuário FENABRAVE (automóveis) × ABVE (leves, todas as tecnologias incl. MHEV) – a série do anuário também inclui micro-híbridos')
hdr(ws,43,['Ano','FENABRAVE – automóveis (P3)','ABVE – leves, com MHEV','Diferença','Diferença %','Fonte da ABVE'],40)
eb=[(2022,'=P3_Eletrificados!D3',49245,'Canal VE (infográfico)'),(2023,'=P3_Eletrificados!D4',93927,'ABVE via Motorshow/Canal VE'),
(2024,'=P3_Eletrificados!D5',177358,'ABVE via Motorshow/Canal VE'),(2025,'=P3_Eletrificados!D6',285300,'Autodata (ABVE, "285,3 mil", arredondado)')]
for i,(a,f,v,s) in enumerate(eb):
    r=44+i  # 44..47
    put(ws,r,1,a,fmt='0'); put(ws,r,2,f,link,UN); put(ws,r,3,v,inp,UN); put(ws,r,4,f"=B{r}-C{r}",fmt=UN); put(ws,r,5,f"=B{r}/C{r}-1",fmt=PC); put(ws,r,6,s)
txt(ws,48,'Diferença de 1% a 2% (o anuário só conta automóveis; a ABVE inclui comerciais leves). Com micro-híbridos fora, a ABVE de 2025 cai para 223.912, e o anuário não tem esse corte.')
widths(ws,[44,16,16,16,16,18,10,10,10,10,14,10,16,16,14,46]); ws.freeze_panes='B3'

# ================= P5_Ranking_Modelos =================
w5=wb.create_sheet('P5_Ranking_Modelos')
title(w5,'P5 – Ranking de modelos de eletrificados – 1º semestre de 2026','Pergunta 3.2.2. Dados da ABVE reproduzidos pelo Canal VE. Tabela em formato longo (pronta para Power BI). O ranking é por versão/modelo como a fonte lista.')
hdr(w5,4,['Pos. na tecnologia','Tecnologia','Modelo (versão)','Marca','Família do modelo','Unidades (1º sem/2026)','% da tecnologia','Fonte'],40)
den={BEV:'P4_Tecnologia!$B$17',PHEV:'P4_Tecnologia!$C$17',HEV:'P4_Tecnologia!$D$17',FLEX:'P4_Tecnologia!$E$17'}
R0=5
for i,(t,pos,mod,mar,fam,u) in enumerate(models):
    r=R0+i
    put(w5,r,1,pos,fmt='0'); put(w5,r,2,t); put(w5,r,3,mod); put(w5,r,4,mar); put(w5,r,5,fam)
    put(w5,r,6,u,inp,UN); put(w5,r,7,f"=F{r}/{den[t]}",fmt=PC); put(w5,r,8,SRC_RANK)
RL=R0+len(models)-1   # 39
rng=lambda c:f"${c}${R0}:${c}${RL}"
# Top 10 geral por versão
r=RL+3; block(w5,r,'A) Top 10 geral – versões/modelos listados, todas as tecnologias (ranking derivado)')
hdr(w5,r+1,['Posição','Modelo (versão)','Marca','Tecnologia','Unidades','% dos eletrificados do semestre'],36)
for k in range(1,11):
    rr=r+1+k
    put(w5,rr,1,k,fmt='0')
    put(w5,rr,5,f"=LARGE({rng('F')},A{rr})",fmt=UN)
    put(w5,rr,2,f"=INDEX({rng('C')},MATCH(E{rr},{rng('F')},0))")
    put(w5,rr,3,f"=INDEX({rng('D')},MATCH(E{rr},{rng('F')},0))")
    put(w5,rr,4,f"=INDEX({rng('B')},MATCH(E{rr},{rng('F')},0))")
    put(w5,rr,6,f"=E{rr}/P4_Tecnologia!$F$17",fmt=PC)
TOP1=r+2; TOP10=r+11
# Top 10 por família
fam={}
for t,pos,mod,mar,f,u in models:
    fam.setdefault(f,[mar,0]); fam[f][1]+=u
top_f=sorted(fam.items(),key=lambda x:-x[1][1])[:10]
r2=TOP10+3; block(w5,r2,'B) Top 10 por família do modelo (versões da mesma família somadas; soma só o que aparece nas listas da fonte)')
hdr(w5,r2+1,['Posição','Família','Marca','Unidades (mín. listado)','% dos eletrificados do semestre'],36)
for k,(f,(mar,_)) in enumerate(top_f,1):
    rr=r2+1+k
    put(w5,rr,1,k,fmt='0'); put(w5,rr,2,f); put(w5,rr,3,mar)
    put(w5,rr,4,f"=SUMIF({rng('E')},B{rr},{rng('F')})",fmt=UN); put(w5,rr,5,f"=D{rr}/P4_Tecnologia!$F$17",fmt=PC)
FAM1=r2+2
txt(w5,r2+12,'A ordem das famílias foi definida na montagem; os valores são fórmulas. Ex.: Haval H6 aparece em 4 listas (PHEV x2, HEV, flex), por isso sobe de posição quando somado.')
# cobertura
r3=r2+15; block(w5,r3,'C) Cobertura das listas (quanto da tecnologia as listas explicam)')
hdr(w5,r3+1,['Tecnologia','Soma listada (un.)','Total da tecnologia (un.)','Cobertura'],32)
for k,(t,ref) in enumerate([(BEV,'B'),(PHEV,'C'),(FLEX,'E'),(HEV,'D')]):
    rr=r3+2+k
    put(w5,rr,1,t); put(w5,rr,2,f"=SUMIF({rng('B')},A{rr},{rng('F')})",fmt=UN)
    put(w5,rr,3,f"=P4_Tecnologia!{ref}17",link,UN); put(w5,rr,4,f"=B{rr}/C{rr}",fmt=PC)
txt(w5,r3+7,'Híbrido flex: a lista de 5 modelos soma 100% do segmento. Os rótulos "versão 19/35" do Haval H6 PHEV seguem a fonte (não confirmei o significado).')
widths(w5,[18,26,38,14,18,18,14,34])
w5.freeze_panes='A5'

# ================= P6_Ranking_Marcas =================
w6=wb.create_sheet('P6_Ranking_Marcas')
title(w6,'P6 – Ranking de marcas de eletrificados','Pergunta 3.2.2. Quatro visões: (A) 1º sem/2026 pela soma dos modelos; (B) elétricos por marca ago–set/2026 (Federação); (C) indicadores da BYD; (D) anuário FENABRAVE 2023–2025.')
block(w6,3,'A) Marcas – soma dos modelos listados em P5 (1º semestre/2026, estimativa)')
hdr(w6,4,['Posição','Marca','BEV (un.)','PHEV (un.)','HEV Flex (un.)','HEV (un.)','Total listado (un.)','% dos eletrificados do semestre'],40)
bt={}
for t,pos,mod,mar,f,u in models: bt[mar]=bt.get(mar,0)+u
brands=sorted(bt,key=lambda x:-bt[x])
for k,mr in enumerate(brands,1):
    r=4+k
    put(w6,r,1,k,fmt='0'); put(w6,r,2,mr,bold if k<=10 else base)
    for c,t in zip((3,4,5,6),(BEV,PHEV,FLEX,HEV)):
        put(w6,r,c,f'=SUMIFS(P5_Ranking_Modelos!{rng("F")},P5_Ranking_Modelos!{rng("D")},$B{r},P5_Ranking_Modelos!{rng("B")},"{t}")',fmt=UN)
    put(w6,r,7,f"=SUM(C{r}:F{r})",fmt=UN,fill=teal_fill if k<=10 else None); put(w6,r,8,f"=G{r}/P4_Tecnologia!$F$17",fmt=PC)
ra=4+len(brands)
txt(w6,ra+1,'Estimativa: só entram modelos que aparecem nas listas da fonte (BEV 90%, PHEV 82% de cobertura). A ordem após o 2º lugar pode mudar com dados completos. A ordem das marcas foi definida na montagem.')

rb=ra+4; block(w6,rb,'B) Elétricos 100% a bateria por marca – Federação (FENABRAVE), via Despachante Dok')
hdr(w6,rb+1,['Marca','Agosto/2026 (un.)','Setembro/2026 (un.)','% set. informado','% set. calculado','Variação set. × ago.','Fonte'],40)
fed=[('BYD',16077,14317,0.649),('Geely',6178,3119,0.141),('GAC',1364,1755,0.08),('GM',571,986,0.045),('MG',646,556,0.025),('Leapmotor',416,400,0.018),('GWM',1286,331,'não informado')]
for i,(m,a,s,p) in enumerate(fed):
    r=rb+2+i; put(w6,r,1,m); put(w6,r,2,a,inp,UN); put(w6,r,3,s,inp,UN)
    put(w6,r,4,p,inp,PC); put(w6,r,5,f"=C{r}/C${rb+10}",fmt=PC); put(w6,r,6,f"=C{r}/B{r}-1",fmt=PC); put(w6,r,7,'Despachante Dok (06/10/2026)')
r=rb+9; put(w6,r,1,'Demais marcas (por diferença)'); put(w6,r,2,f"=B{rb+10}-SUM(B{rb+2}:B{rb+8})",fmt=UN); put(w6,r,3,f"=C{rb+10}-SUM(C{rb+2}:C{rb+8})",fmt=UN)
put(w6,r,4,''); put(w6,r,5,f"=C{r}/C${rb+10}",fmt=PC); put(w6,r,6,f"=C{r}/B{r}-1",fmt=PC); put(w6,r,7,'Calculado')
r=rb+10; put(w6,r,1,'TOTAL de elétricos',bold); put(w6,r,2,26989,inp,UN); put(w6,r,3,22007,inp,UN); put(w6,r,4,''); put(w6,r,5,f"=SUM(E{rb+2}:E{rb+9})",fmt=PC); put(w6,r,6,f"=C{r}/B{r}-1",fmt=PC); put(w6,r,7,'Despachante Dok (06/10/2026)')
txt(w6,rb+11,'A Federação informa 64,9% para a BYD; 14.317 ÷ 22.007 = 65,1% (denominador não divulgado). Total da Federação em set.: 22.007; ABVE: 22.246 (diferença de 239).')

rc=rb+14; block(w6,rc,'C) Indicadores da BYD e do mercado de elétricos (Federação / ABVE)')
hdr(w6,rc+1,['Indicador','Valor','Fonte'],30)
ind=[('Participação da BYD nos elétricos – jan. a set./2026',0.629,PC,'Despachante Dok'),
('Elétricos vendidos jan. a set./2026 (aprox.)',164000,UN,'Despachante Dok'),
('Elétricos vendidos jan. a set./2025',52800,UN,'Despachante Dok'),
('Crescimento 2026 sobre 2025',f"=B{rc+3}/B{rc+4}-1",PC,'Calculado'),
('BYD – elétricos acumulados até ago./2026',88933,UN,'Gazeta de São Paulo (Federação)'),
('BYD – híbridos acumulados até ago./2026',57541,UN,'Gazeta de São Paulo (Federação)'),
('BYD – total de eletrificados até ago./2026',f"=B{rc+6}+B{rc+7}",UN,'Calculado')]
for i,(a,v,fm,s) in enumerate(ind):
    r=rc+2+i; put(w6,r,1,a); put(w6,r,2,v,base if isinstance(v,str) else inp,fm); put(w6,r,3,s)

rd=rc+11; block(w6,rd,'D) Anuário FENABRAVE – elétricos por marca (Top 15, ano cheio, automóveis)')
hdr(w6,rd+1,['Marca','2023 (un.)','2024 (un.)','2025 (un.)','% do total BEV 2025','Cresc. 2025 × 2024'],36)
for i,(m,a,b,c) in enumerate(brand_bev):
    r=rd+2+i; put(w6,r,1,m,bold if i<3 else base)
    for col,v in zip((2,3,4),(a,b,c)): put(w6,r,col,v if v is not None else 0,inp,UN)
    put(w6,r,5,f"=D{r}/P3_Eletrificados!$B$6",fmt=PC); put(w6,r,6,f'=IF(C{r}>0,D{r}/C{r}-1,"n/d")',fmt=PC)
rD1=rd+2; rD2=rd+16
put(w6,rD2+1,1,'Soma do Top 15',bold)
for col in (2,3,4): put(w6,rD2+1,col,f"=SUM({L(col)}{rD1}:{L(col)}{rD2})",bold,UN)
put(w6,rD2+1,5,f"=D{rD2+1}/P3_Eletrificados!$B$6",bold,PC); put(w6,rD2+1,6,'')
txt(w6,rD2+2,'O anuário diz "BYD 74,9% dos elétricos em 2025"; sobre o total de 79.351 do próprio gráfico dá 71,9%. Mantido o cálculo sobre o total. "–" no anuário foi tratado como 0.')

re_=rD2+5; block(w6,re_,'E) Anuário FENABRAVE – híbridos por marca (Top 15, ano cheio; inclui micro-híbridos)')
hdr(w6,re_+1,['Marca','2023 (un.)','2024 (un.)','2025 (un.)','% do total híbridos 2025','Cresc. 2025 × 2024'],36)
for i,(m,a,b,c) in enumerate(brand_hyb):
    r=re_+2+i; put(w6,r,1,m,bold if i<3 else base)
    for col,v in zip((2,3,4),(a,b,c)): put(w6,r,col,v if v is not None else 0,inp,UN)
    put(w6,r,5,f"=D{r}/P3_Eletrificados!$C$6",fmt=PC); put(w6,r,6,f'=IF(C{r}>0,D{r}/C{r}-1,"n/d")',fmt=PC)
rE2=re_+16
put(w6,rE2+1,1,'Soma do Top 15',bold)
for col in (2,3,4): put(w6,rE2+1,col,f"=SUM({L(col)}{re_+2}:{L(col)}{rE2})",bold,UN)
put(w6,rE2+1,5,f"=D{rE2+1}/P3_Eletrificados!$C$6",bold,PC); put(w6,rE2+1,6,'')
txt(w6,rE2+2,'Fiat (42.601 em 2025) e parte da Mercedes-Benz são, em geral, micro-híbridos: por isso a Fiat é 2ª em híbridos no anuário e não aparece no ranking da ABVE (sem MHEV). Interpretação minha, coerente com o bloco E de P4.')
widths(w6,[46,18,18,18,20,18,34,22])
w6.freeze_panes='A3'

# ================= P7_Geografia =================
w7=wb.create_sheet('P7_Geografia')
title(w7,'P7 – Distribuição geográfica dos emplacamentos de eletrificados','Pergunta 3.2.3. Dados da ABVE ("Geografia da Eletromobilidade"), via comunicados e reportagens. Emplacamentos (vendas), não frota. Frota está em P8.')
block(w7,3,'A) Regiões')
hdr(w7,4,['Região','1º sem/2026 (un.)','% do Brasil','Ano 2025 (un.)','% do Brasil 2025','Diferença (pp)'],36)
regs=[('Sudeste',96159,103964),('Nordeste',40310,36596),('Sul',38716,None),('Centro-Oeste',30426,33964),('Norte',9411,9303)]
for i,(n,a,b) in enumerate(regs):
    r=5+i; put(w7,r,1,n,bold if i==0 else base); put(w7,r,2,a,inp,UN); put(w7,r,3,f"=B{r}/P4_Tecnologia!$F$17",fmt=PC)
    if b is None: put(w7,r,4,'=P4_Tecnologia!$M$16-SUM(D5:D6,D8:D9)',fmt=UN,fill=teal_fill)
    else: put(w7,r,4,b,inp,UN)
    put(w7,r,5,f"=D{r}/P4_Tecnologia!$M$16",fmt=PC); put(w7,r,6,f"=(C{r}-E{r})*100",fmt='0.0')
put(w7,10,1,'Soma das regiões',bold); put(w7,10,2,'=SUM(B5:B9)',bold,UN); put(w7,10,3,'=SUM(C5:C9)',bold,PC); put(w7,10,4,'=SUM(D5:D9)',bold,UN); put(w7,10,5,'=SUM(E5:E9)',bold,PC); put(w7,10,6,'')
txt(w7,11,'Sul/2025 (célula verde-água) estava ausente na planilha original: calculado por diferença do total ABVE de 2025 (223.912) = 40.085, ou 17,9% – confere com o 17,9% divulgado (Vrum). 1º sem/2026: soma 215.022 × total 215.023 (1 unidade de arredondamento da fonte).')

block(w7,13,'B) Estados – cinco maiores')
hdr(w7,14,['Estado','UF','Região','1º sem/2026 (un.)','% do Brasil','Ano 2025 (un.)','% do Brasil 2025'],36)
sts=[('São Paulo','SP','Sudeste',61629,68618),('Distrito Federal','DF','Centro-Oeste',18131,21639),('Minas Gerais','MG','Sudeste',16341,15155),('Paraná','PR','Sul',14738,14024),('Rio de Janeiro','RJ','Sudeste',12959,14280)]
for i,(n,uf,rg,a,b) in enumerate(sts):
    r=15+i; put(w7,r,1,n,bold if i==0 else base); put(w7,r,2,uf); put(w7,r,3,rg); put(w7,r,4,a,inp,UN); put(w7,r,5,f"=D{r}/P4_Tecnologia!$F$17",fmt=PC)
    put(w7,r,6,b,inp,UN); put(w7,r,7,f"=F{r}/P4_Tecnologia!$M$16",fmt=PC)
put(w7,20,1,'Soma dos 5 maiores',bold); put(w7,20,2,''); put(w7,20,3,''); put(w7,20,4,'=SUM(D15:D19)',bold,UN); put(w7,20,5,'=SUM(E15:E19)',bold,PC); put(w7,20,6,'=SUM(F15:F19)',bold,UN); put(w7,20,7,'=SUM(G15:G19)',bold,PC)
txt(w7,21,'2025: o Paraná (14.024; 6,3%) não estava na planilha original e foi acrescentado (Forbes Brasil/ABVE). Note que RJ e MG trocam de posição entre 2025 e 1º sem/2026.')

block(w7,23,'C) Cidades – cinco maiores (1º semestre/2026)')
hdr(w7,24,['Cidade','UF','Região','1º sem/2026 (un.)','% do Brasil'],30)
cts=[('São Paulo','SP','Sudeste',22566),('Brasília','DF','Centro-Oeste',18131),('Belo Horizonte','MG','Sudeste',9896),('Rio de Janeiro','RJ','Sudeste',7096),('Curitiba','PR','Sul',6395)]
for i,(n,uf,rg,a) in enumerate(cts):
    r=25+i; put(w7,r,1,n); put(w7,r,2,uf); put(w7,r,3,rg); put(w7,r,4,a,inp,UN); put(w7,r,5,f"=D{r}/P4_Tecnologia!$F$17",fmt=PC)
put(w7,30,1,'Soma das 5 cidades',bold); put(w7,30,2,''); put(w7,30,3,''); put(w7,30,4,'=SUM(D25:D29)',bold,UN); put(w7,30,5,'=SUM(E25:E29)',bold,PC)
txt(w7,31,'Brasília = Distrito Federal: os 18.131 da cidade e do estado são o mesmo número (tudo no DF).')

block(w7,33,'D) Setembro/2026 (total do mês: 53.126)')
hdr(w7,34,['Local','UF','Nível','Emplacamentos (un.)','% do Brasil','Observação'],30)
sep=[('São Paulo','SP','Estado','=E35*P4_Tecnologia!$F$19',0.268,'Unidades estimadas (percentual informado × total)'),
('Minas Gerais','MG','Estado','não informado','não informado','2º lugar, sem valor na fonte'),
('Distrito Federal','DF','Estado',4158,None,''),('Paraná','PR','Estado',3403,None,''),('Rio de Janeiro','RJ','Estado',3166,None,''),
('São Paulo','SP','Cidade',5240,None,''),('Brasília','DF','Cidade',4158,None,''),('Belo Horizonte','MG','Cidade',2451,None,''),('Rio de Janeiro','RJ','Cidade',1625,None,''),('Curitiba','PR','Cidade',1467,None,'')]
for i,(n,uf,nv,u,p,ob) in enumerate(sep):
    r=35+i; put(w7,r,1,n); put(w7,r,2,uf); put(w7,r,3,nv)
    if isinstance(u,str) and u.startswith('='): put(w7,r,4,u,fmt=UN)
    else: put(w7,r,4,u,inp,UN)
    if p is None: put(w7,r,5,f"=D{r}/P4_Tecnologia!$F$19",fmt=PC)
    else: put(w7,r,5,p,inp,PC)
    put(w7,r,6,ob)
# fix: SP estimado usa % da própria linha
w7['D35']='=E35*P4_Tecnologia!$F$19'

block(w7,46,'E) Concentração (1º semestre/2026)')
hdr(w7,47,['Indicador','Valor','Leitura'],30)
conc=[('Sudeste no Brasil','=C5','Mais de 4 em cada 10 emplacamentos'),('Nordeste + Norte','=C6+C9','Regiões com menor participação juntas'),
('São Paulo (estado) no Brasil','=E15','Maior estado isolado'),('5 maiores estados no Brasil','=E20','Concentração em poucos estados'),
('Distrito Federal no Brasil','=E16','2º lugar com população pequena: forte incentivo e frota corporativa'),('Mudança do Sudeste: 1º sem/2026 − 2025 (pp)','=F5','Negativo = desconcentração')]
for i,(a,v,l) in enumerate(conc):
    r=48+i; put(w7,r,1,a); put(w7,r,2,v,fmt=('0.0' if 'pp' in a else PC)); put(w7,r,3,l)
widths(w7,[42,16,16,18,16,48,16]); w7.freeze_panes='A3'

# ================= P8_Frota_Ref =================
w8=wb.create_sheet('P8_Frota_Ref')
title(w8,'P8 – Frota circulante de eletrificados (referência – NÃO é venda do ano)','Fonte: Secretaria Nacional de Trânsito (SENATRAN), via levantamentos e reportagens; planilhas oficiais da SENATRAN não puderam ser abertas neste ambiente.')
block(w8,3,'A) Frota nacional – dezembro/2025 (NeoCharge com dados da SENATRAN)')
hdr(w8,4,['Tecnologia','Veículos','% da frota','Fonte'],28)
for i,(n,v) in enumerate([('Elétrico 100% a bateria',154085),('Híbrido plug-in',200244),('Híbrido convencional',259060)]):
    r=5+i; put(w8,r,1,n); put(w8,r,2,v,inp,UN); put(w8,r,3,f"=B{r}/$B$8",fmt=PC); put(w8,r,4,'Cenário Energia (02/03/2026)')
put(w8,8,1,'TOTAL',bold); put(w8,8,2,'=SUM(B5:B7)',bold,UN); put(w8,8,3,'=SUM(C5:C7)',bold,PC); put(w8,8,4,'Calculado')
put(w8,9,1,'Total informado pela reportagem'); put(w8,9,2,613389,inp,UN); put(w8,9,3,''); put(w8,9,4,'Cenário Energia')
put(w8,10,1,'Diferença (deve ser zero)'); put(w8,10,2,'=B8-B9',fmt=UN); put(w8,10,3,''); put(w8,10,4,'Calculado')

block(w8,12,'B) São Paulo – dezembro/2025')
hdr(w8,13,['Recorte','Veículos','% do estado','Fonte'],28)
put(w8,14,1,'Estado de São Paulo'); put(w8,14,2,196344,inp,UN); put(w8,14,3,'=B14/B14',fmt=PC); put(w8,14,4,'Motorshow (NeoCharge/SENATRAN)')
put(w8,15,1,'Cidade de São Paulo'); put(w8,15,2,86681,inp,UN); put(w8,15,3,'=B15/B14',fmt=PC); put(w8,15,4,'Motorshow (NeoCharge/SENATRAN)')

block(w8,17,'C) Frota por estado – ABVE, janeiro de 2022 a janeiro de 2026 (só veículos emplacados desde 2022)')
hdr(w8,18,['Estado','Veículos','','Fonte'],28)
for i,(n,v) in enumerate([('São Paulo',181305),('Distrito Federal',48502),('Rio de Janeiro',39295),('Minas Gerais',37897),('Paraná',37703)]):
    r=19+i; put(w8,r,1,n); put(w8,r,2,v,inp,UN); put(w8,r,3,''); put(w8,r,4,'Canal VE (ABVE)')
txt(w8,24,'Atenção: a frota da ABVE (desde 2022) não é comparável à frota total da SENATRAN (ex.: SP 181.305 × 196.344).')

block(w8,26,'D) Distrito Federal – DETRAN-DF (setembro/2026)')
hdr(w8,27,['Indicador','Valor','','Fonte'],28)
put(w8,28,1,'Elétricos 100% a bateria'); put(w8,28,2,19659,inp,UN); put(w8,28,3,''); put(w8,28,4,'Correio da Manhã (DETRAN-DF)')
put(w8,29,1,'Híbridos'); put(w8,29,2,39527,inp,UN); put(w8,29,3,''); put(w8,29,4,'Correio da Manhã (DETRAN-DF)')
put(w8,30,1,'Total de eletrificados'); put(w8,30,2,'=B28+B29',fmt=UN); put(w8,30,3,''); put(w8,30,4,'Calculado')
put(w8,31,1,'Frota total do Distrito Federal'); put(w8,31,2,2223419,inp,UN); put(w8,31,3,''); put(w8,31,4,'Correio da Manhã (DETRAN-DF)')
put(w8,32,1,'Eletrificados na frota total'); put(w8,32,2,'=B30/B31',fmt=PC); put(w8,32,3,''); put(w8,32,4,'Calculado')

block(w8,34,'E) Vendas anuais de eletrificados – ABVE (atenção ao critério de cada ano)')
hdr(w8,35,['Ano','Emplacamentos','Critério','Fonte'],28)
ev=[('2012',117,'sem informação','Gazeta de São Paulo (ABVE)'),('2016',1091,'sem informação','Gazeta de São Paulo (ABVE)'),('2020',19745,'sem informação','Gazeta de São Paulo (ABVE)'),
('2024','=P4_Tecnologia!M15','COM micro-híbridos (MHEV)','Gazeta de São Paulo / Motorshow'),('2025','=P4_Tecnologia!M16','SEM micro-híbridos (critério vigente desde 2025)','Gazeta de São Paulo / Canal VE'),
('2026 (jan.–set.)',381609,'SEM micro-híbridos','Gazeta de São Paulo (ABVE)'),('Acumulado desde 2012 (informado)',1003310,'critérios misturados','Gazeta de São Paulo (ABVE)')]
for i,(a,v,cr,s) in enumerate(ev):
    r=36+i; put(w8,r,1,a); put(w8,r,2,v,base if isinstance(v,str) else inp,UN); put(w8,r,3,cr); put(w8,r,4,s)
put(w8,44,1,'Crescimento 2025 × 2024 como divulgado (critérios misturados)'); put(w8,44,2,'=B40/B39-1',fmt=PC); put(w8,44,3,'≈ +26% (divulgado pela ABVE)'); put(w8,44,4,'Calculado')
put(w8,45,1,'Crescimento 2025 × 2024 no mesmo critério (sem MHEV)',bold); put(w8,45,2,'=P4_Tecnologia!F16/P4_Tecnologia!F15-1',bold,PC,teal_fill); put(w8,45,3,'2024 sem MHEV = 161.172'); put(w8,45,4,'Calculado (P4, bloco B)')
put(w8,46,1,'Conferência jan.–set./2026: 1º sem + jul + ago + set'); put(w8,46,2,'=P4_Tecnologia!F17+56074+57386+P4_Tecnologia!F19',fmt=UN); put(w8,46,3,'confere com 381.609'); put(w8,46,4,'Jul: Paraíba Business; Ago: Revista Amazônia/Autopapo')
widths(w8,[58,16,46,36])
for ws_ in (w8,): ws_.freeze_panes='A3'
wb.save(OUT); print('saved part1')
