from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
SRC='/mnt/user-data/outputs/Projeto_BI_Automotivo_Etapa2_Eletrificados.xlsx'
OUT='/mnt/user-data/outputs/Projeto_BI_Automotivo_Etapa3_Renda_Personas.xlsx'
wb=load_workbook(SRC)
F='Arial'; NAVY='1B365D'; TEAL='008080'
hdr_fill=PatternFill('solid',fgColor=NAVY); hdr_font=Font(name=F,bold=True,color='FFFFFF',size=10)
teal_fill=PatternFill('solid',fgColor='D9EEEE')
base=Font(name=F,size=10); bold=Font(name=F,size=10,bold=True); inp=Font(name=F,size=10,color='0000FF'); link=Font(name=F,size=10,color='008000')
note=Font(name=F,size=9,italic=True,color='555555')
thin=Side(style='thin',color='BFC5CF'); box=Border(top=thin,bottom=thin,left=thin,right=thin)
UN='#,##0;-#,##0;0'; PC='0.0%'; RS='"R$ "#,##0'
def title(ws,t,sub=None):
    ws['A1']=t; ws['A1'].font=Font(name=F,size=14,bold=True,color=NAVY)
    if sub: ws['A2']=sub; ws['A2'].font=note
    ws.sheet_view.showGridLines=False
def block(ws,r,t): ws.cell(r,1,t).font=Font(name=F,size=11,bold=True,color=TEAL)
def hdr(ws,r,labels,h=48):
    for i,l in enumerate(labels,1):
        c=ws.cell(r,i,l); c.fill=hdr_fill; c.font=hdr_font; c.border=box; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[r].height=h
def put(ws,r,c,v,font=base,fmt=None,fill=None,wrap=True):
    x=ws.cell(r,c,v); x.font=font; x.border=box; x.alignment=Alignment(wrap_text=wrap,vertical='top')
    if fmt: x.number_format=fmt
    if fill: x.fill=fill
    return x
def txt(ws,r,t,col=1): c=ws.cell(r,col,t); c.font=note
def widths(ws,w):
    for i,x in enumerate(w,1): ws.column_dimensions[L(i)].width=x

# ============ P9_Faixas_Renda ============
w=wb.create_sheet('P9_Faixas_Renda')
title(w,'P9 – Renda familiar, regra dos 20% e faixa de preço recomendada','Pergunta 3.3.1. Une a distribuição de renda do IBGE (POF 2017-2018) às 5 faixas do escopo e calcula quanto cada faixa consegue pagar. Células azuis são parâmetros que podem ser alterados.')
block(w,3,'A) Parâmetros')
hdr(w,4,['Parâmetro','Valor','Fonte / observação'],28)
params=[('Salário mínimo 2026 (R$)',1621,RS,'Decreto 12.797/2025 (vigente desde 1º/jan/2026)'),
('Regra: % da renda mensal para a parcela',0.20,'0%','Regra dos 20% do escopo do projeto'),
('Juros do financiamento (% ao mês)',0.0204,'0.00%','BCB: média das novas operações de crédito para veículos, 2,04% a.m. (27,4% a.a.), citada em jun/2026. Atualizar na série 20749 do SGS/BCB'),
('Prazo do financiamento (meses)',48,'0','Premissa do projeto (48 parcelas)'),
('Entrada (% do preço)',0.20,'0%','Premissa do projeto. Altere para ver o efeito'),
('Salário mínimo 2018 (R$)',954,RS,'Base da POF 2017-2018 (2 SM = R$ 1.908)'),
('Transporte na despesa de consumo – média nacional POF 2017-2018',0.181,'0.0%','IBGE: alimentação 17,5% e transporte 0,6 p.p. acima. Só referência de contexto')]
for i,(a,v,fm,s) in enumerate(params):
    r=5+i; put(w,r,1,a); put(w,r,2,v,inp,fm); put(w,r,3,s)

block(w,13,'B) Distribuição das famílias brasileiras por classe de renda – POF 2017-2018 (IBGE)')
hdr(w,14,['Classe de rendimento familiar','Limite inferior (SM)','Limite superior (SM)','Inferior (R$ 2018)','Superior (R$ 2018)','Inferior (R$ 2026, mesmo nº de SM)','Superior (R$ 2026, mesmo nº de SM)','Famílias (milhões)','% das famílias','% acumulado','Fonte'],56)
cls=[('Até 2 SM',0,2,16.4,0.239),('Mais de 2 a 3 SM',2,3,12.8,None),('Mais de 3 a 6 SM',3,6,21.8,None),('Mais de 6 a 10 SM',6,10,9.6,None),('Mais de 10 a 15 SM',10,15,4.4,None),('Mais de 15 a 25 SM',15,25,None,None),('Mais de 25 SM',25,None,None,0.027)]
for i,(n,lo,hi,fam,pct) in enumerate(cls):
    r=15+i
    put(w,r,1,n,bold); put(w,r,2,lo,inp,'0'); put(w,r,3,hi if hi else 'sem limite',inp if hi else base,'0')
    put(w,r,4,f"=B{r}*$B$10",fmt=RS); put(w,r,5,f"=C{r}*$B$10" if hi else 'sem limite',fmt=RS)
    put(w,r,6,f"=B{r}*$B$5",fmt=RS); put(w,r,7,f"=C{r}*$B$5" if hi else 'sem limite',fmt=RS)
    if i==5: put(w,r,8,"=H22-SUM(H15:H19)-H21",fmt='0.0',fill=teal_fill)
    elif i==6: put(w,r,8,"=I21*H22",fmt='0.0')
    else: put(w,r,8,fam,inp,'0.0')
    if pct is not None: put(w,r,9,pct,inp,PC)
    else: put(w,r,9,f"=H{r}/$H$22",fmt=PC,fill=teal_fill if i==5 else None)
    put(w,r,10,f"=SUM($I$15:I{r})",fmt=PC)
    put(w,r,11,'IBGE, POF 2017-2018 (release e matérias que o reproduzem)' if i not in (5,) else 'Calculada por diferença (valor não localizado)')
put(w,22,1,'Total de famílias',bold); 
for c in range(2,8): put(w,22,c,'')
put(w,22,8,"=H15/I15",bold,'0.0'); put(w,22,9,"=SUM(I15:I21)",bold,PC); put(w,22,10,''); put(w,22,11,'Total derivado de 16,4 milhões = 23,9% das famílias')
txt(w,23,'Fontes dos números: 23,9% (16,4 milhões) até 2 SM e 2,7% acima de 25 SM no release do IBGE; 12,8, 21,8, 9,6 e 4,4 milhões nas classes intermediárias segundo a Carta Capital. A classe de 15 a 25 SM é calculada por diferença.')
txt(w,24,'A POF 2017-2018 é a edição mais recente com resultados publicados que encontrei. Os valores em R$ de 2018 foram convertidos por número de salários mínimos, e não por inflação.')

block(w,26,'C) As 5 faixas do escopo: classe POF e capacidade de compra pela regra dos 20%')
hdr(w,27,['Faixa','Persona','Renda mensal mín. (R$)','Renda mensal máx. (R$)','Renda média (R$)','Renda média em SM 2026','Classe POF (nº)','Classe POF','% das famílias na classe','% das famílias até o fim da classe','Parcela máx. (renda mín.)','Parcela máx. (renda máx.)','Valor financiável – mín.','Valor financiável – máx.','Preço máx. pela regra – mín.','Preço máx. pela regra – máx.','Teto do escopo – mín.','Teto do escopo – máx.','Entrada necessária p/ o teto máx. do escopo','Avaliação do teto do escopo'],70)
pers=[(1,'Seu Zé / Cláudio (Mobilidade essencial)',1800,3000,12000,25000),(2,'Ana Paula (Entrada / seminovo)',3500,5500,0,65000),(3,'Lucas & Mariana (Classe média)',9000,15000,85000,130000),(4,'Roberto (Executivo)',20000,35000,160000,240000),(5,'Dra. Helena (Alta renda)',45000,45000,300000,300000)]
for i,(fx,pn,a,b,t0,t1) in enumerate(pers):
    r=28+i
    put(w,r,1,fx,fmt='0'); put(w,r,2,pn,bold); put(w,r,3,a,inp,RS); put(w,r,4,b,inp,RS)
    put(w,r,5,f"=AVERAGE(C{r}:D{r})",fmt=RS); put(w,r,6,f"=E{r}/$B$5",fmt='0.0')
    put(w,r,7,f"=MATCH(F{r},$B$15:$B$21,1)",fmt='0'); put(w,r,8,f"=INDEX($A$15:$A$21,G{r})")
    put(w,r,9,f"=INDEX($I$15:$I$21,G{r})",fmt=PC); put(w,r,10,f"=INDEX($J$15:$J$21,G{r})",fmt=PC)
    put(w,r,11,f"=C{r}*$B$6",fmt=RS); put(w,r,12,f"=D{r}*$B$6",fmt=RS)
    put(w,r,13,f"=PV($B$7,$B$8,-K{r})",fmt=RS); put(w,r,14,f"=PV($B$7,$B$8,-L{r})",fmt=RS)
    put(w,r,15,f"=M{r}/(1-$B$9)",fmt=RS,fill=teal_fill); put(w,r,16,f"=N{r}/(1-$B$9)",fmt=RS,fill=teal_fill)
    put(w,r,17,t0,inp,RS); put(w,r,18,t1,inp,RS)
    put(w,r,19,f"=MAX(0,1-N{r}/R{r})",fmt=PC)
    put(w,r,20,f'=IF(R{r}<=P{r},"Teto cabe na regra com entrada de "&TEXT($B$9,"0%"),"Teto exige entrada de "&TEXT(S{r},"0%")&" ou prazo/consórcio maior")')
    w.row_dimensions[r].height=40
txt(w,33,'Dra. Helena: renda "acima de R$ 45 mil" e teto "acima de R$ 300 mil" entram como valores mínimos; a capacidade real é maior. Ana Paula: o escopo não define piso de orçamento (0 = sem piso).')
txt(w,34,'Fórmula: valor financiável = VP(juros, prazo, −parcela); preço máximo = financiável ÷ (1 − entrada). Não inclui IPVA, seguro, combustível nem manutenção.')

block(w,36,'D) Faixa de preço e categoria recomendadas por nível de renda (resposta à pergunta 3.3.1)')
hdr(w,37,['Faixa','Persona','Preço mínimo pela regra (R$)','Preço máximo pela regra (R$)','Categoria recomendada','Eletrificado?','Justificativa'],36)
rec=[('Usado antigo (popular 1.0, 2004–2010)','Inviável','Cabe só um popular antigo: capacidade de R$ 14 mil a R$ 23 mil. O elétrico mais barato (R$ 119.990) custa quase 5 vezes o teto de R$ 25 mil.'),
('Usado/seminovo de entrada','Só em transição (sem preço FIPE de elétrico usado localizado)','A regra dá R$ 27 mil a R$ 42 mil. Onix e HB20 2019 custam R$ 49 mil a R$ 54 mil e exigem entrada de 32% a 38% mesmo na renda máxima da faixa (57% a 61% na renda mínima).'),
('Zero-km de entrada ou SUV seminovo; elétrico de entrada só na renda alta da faixa','Possível (Dolphin Mini)','Capacidade de R$ 68 mil a R$ 114 mil. Polo Track (R$ 98 mil) e Tracker 2022/23 (R$ 89 mil a R$ 91 mil) só ficam abaixo de 20% na parte alta da faixa (renda acima de ~R$ 12 mil a R$ 13 mil); o Dolphin Mini (R$ 120 mil) passa de 20% mesmo no topo (21%).'),
('SUV novo ou híbrido (flex ou plug-in)','Sim (Haval H6, Song Plus, Corolla Cross híbrido)','Capacidade de R$ 152 mil a R$ 266 mil. O Corolla Cross híbrido seminovo 2022 (R$ 139 mil) cabe até na renda mínima; T-Cross, Corolla Cross 0 km e Haval H6 só na parte alta; o Song Plus (cerca de R$ 250 mil) passa o teto de R$ 240 mil do escopo.'),
('Premium a gasolina ou elétrico premium','Sim (BYD Seal e similares)','Capacidade a partir de R$ 342 mil, acima do piso de R$ 300 mil do escopo. O BYD Seal (R$ 299.990) cabe com folga.')]
for i,(pn,cat,el,just) in [(i,(p[1],)+r_) for i,(p,r_) in enumerate(zip(pers,rec))]:
    pass
for i,(p,(cat,el,just)) in enumerate(zip(pers,rec)):
    r=38+i
    put(w,r,1,p[0],fmt='0'); put(w,r,2,p[1],bold)
    put(w,r,3,f"=ROUND(O{28+i},-3)",fmt=RS); put(w,r,4,f"=ROUND(P{28+i},-3)",fmt=RS)
    put(w,r,5,cat); put(w,r,6,el); put(w,r,7,just); w.row_dimensions[r].height=58
txt(w,43,'Dra. Helena: o valor da coluna "máximo" é o da renda mínima de R$ 45 mil (a capacidade cresce com a renda). Faixas de preço calculadas com os parâmetros do bloco A.')
widths(w,[34,34,14,14,14,12,10,22,12,13,14,14,15,15,15,15,14,14,16,46]); w.freeze_panes='A3'

# ============ P10_Catalogo_Precos ============
c=wb.create_sheet('P10_Catalogo_Precos')
title(c,'P10 – Catálogo de preços de referência (zero-km e Tabela FIPE)','Pergunta 3.3.2. Usado = valor médio da Tabela FIPE (mês indicado). Zero-km = preço de tabela sugerido, sem descontos de loja. Aba em formato longo, pronta para Power BI.')
hdr(c,4,['ID','Veículo','Ano-modelo','Categoria','Tecnologia','Preço (R$)','Tipo de preço','Referência','Fonte','Situação da verificação'],40)
A='Conferido nesta rodada'; B='Do arquivo recebido; não reverificado'; P='Aproximado (reportagem cita "cerca de")'
cat=[
('Fiat Uno Mille Celeb 1.0 Flex 2p','2008','Usado antigo','Combustão',16821,'Tabela FIPE','Out/2026','tabelafipebrasil.com (reproduz a FIPE)',A),
('Fiat Uno Mille Fire/Economy 1.0 4p','2008','Usado antigo','Combustão',21674,'Tabela FIPE','Out/2026','tabelafipebrasil.com',A),
('Fiat Palio 1.0 Fire Flex 2p','2008','Usado antigo','Combustão',22677,'Tabela FIPE','Ago/2026','tabelafipebrasil.com',A),
('Fiat Palio ELX 1.4 Fire Flex 4p','2008','Usado antigo','Combustão',28264,'Tabela FIPE','Ago/2026','tabelafipebrasil.com',A),
('Chevrolet Corsa Hatch Joy 1.0 Flex','2008','Usado antigo','Combustão',21200,'Tabela FIPE','Jul/2026','tabelafipebrasil.com',A),
('Chevrolet Celta 1.0 3p','2004','Usado antigo','Combustão',13755,'Tabela FIPE','Out/2026','tabelafipebrasil.com (ano 2004; não achei 2008)',A),
('Chevrolet Onix Hatch Joy 1.0','2019','Usado/seminovo','Combustão',49326,'Tabela FIPE','Out/2026','tabelafipebrasil.com',A),
('Chevrolet Onix Hatch LT 1.0','2019','Usado/seminovo','Combustão',53269,'Tabela FIPE','Out/2026','tabelafipebrasil.com',A),
('Chevrolet Onix Hatch Activ 1.4 Aut.','2019','Usado/seminovo','Combustão',67619,'Tabela FIPE','Out/2026','tabelafipebrasil.com',A),
('Hyundai HB20S Unique 1.0 Mec.','2019','Usado/seminovo','Combustão',54021,'Tabela FIPE','Jul/2026','carzin.com.br (reproduz a FIPE)',A),
('Hyundai HB20S Style 1.6 Aut.','2019','Usado/seminovo','Combustão',68345,'Tabela FIPE','Jul/2026','carzin.com.br',A),
('Chevrolet Tracker LT 1.0 Turbo Aut.','2022','Seminovo','Combustão',88933,'Tabela FIPE','Out/2026','tabelafipebrasil.com',A),
('Chevrolet Tracker LT 1.0 Turbo Aut.','2023','Seminovo','Combustão',91308,'Tabela FIPE','Out/2026','tabelafipebrasil.com',A),
('Chevrolet Tracker Premier 1.2 Turbo Aut.','2023','Seminovo','Combustão',104475,'Tabela FIPE','Set/2026','tabelafipebrasil.com',A),
('Toyota Corolla Cross XR 2.0 Flex Aut.','2022','Seminovo','Combustão',123776,'Tabela FIPE','Set/2026','tabelafipebrasil.com',A),
('Toyota Corolla Cross XRX 1.8 Híbrido','2022','Seminovo','Híbrido (HEV)',139427,'Tabela FIPE','Set/2026','tabelafipebrasil.com',A),
('Fiat Mobi Like','2026','Zero-km','Combustão',72490,'Tabela / oferta oficial','Nov/2026','Bright Consulting / UOL Carros (arquivo recebido)',B),
('Citroën C3 Live','2026','Zero-km','Combustão',77990,'Tabela / oferta oficial','Nov/2026','Bright Consulting / UOL Carros (arquivo recebido)',B),
('Renault Kwid Evolution','2026','Zero-km','Combustão',82790,'Tabela / oferta oficial','Nov/2026','Bright Consulting / UOL Carros (arquivo recebido)',B),
('Hyundai HB20 Comfort','2026','Zero-km','Combustão',96140,'Tabela / oferta oficial','Nov/2026','Bright Consulting / UOL Carros (arquivo recebido)',B),
('Volkswagen Polo Track 1.0','2026','Zero-km','Combustão',97990,'Tabela / oferta oficial','Nov/2026','Bright Consulting / UOL Carros (arquivo recebido); Localiza cita R$ 96.690',B),
('Fiat Argo','2026','Zero-km','Combustão',98990,'Tabela / oferta oficial','Nov/2026','Bright Consulting / UOL Carros (arquivo recebido)',B),
('BYD Dolphin Mini','2026','Zero-km','Elétrico (BEV)',119990,'Tabela / oferta oficial','Nov/2026','Auto+ / WM1 (arquivo recebido)',B),
('Geely EX2','2026','Zero-km','Elétrico (BEV)',123800,'Tabela','Nov/2026','iCarros (arquivo recebido)',B),
('Volkswagen T-Cross 200 TSI','2026','Zero-km','Combustão',151490,'Tabela sugerida','Jul/2026','Tribuna PR (corte de preços da VW)',A),
('Volkswagen T-Cross Comfortline','2026','Zero-km','Combustão',171990,'Tabela sugerida','Jul/2026','Tribuna PR (corte de preços da VW)',A),
('Volkswagen T-Cross Highline','2026','Zero-km','Combustão',186290,'Tabela sugerida','Set/2026','Autoo (comparação de preços de 0 km)',A),
('Chevrolet Tracker Premier','2026','Zero-km','Combustão',177990,'Tabela sugerida','Set/2026','Autoo (comparação de preços de 0 km)',A),
('Toyota Corolla Cross XRE (estoque 2026, em oferta)','2026','Zero-km','Combustão',179590,'Preço promocional','Jul/2026','Canaltech (09/07/2026)',A),
('Toyota Corolla Cross XRE (linha 2027, tabela)','2027','Zero-km','Combustão',197120,'Tabela','Jul/2026','Canaltech (09/07/2026)',A),
('GWM Haval H6 (a partir de)','2026','Zero-km','Híbrido (PHEV/HEV)',239000,'Tabela (aprox.)','2026 (mês não informado)','Motor Agora (cita "cerca de R$ 239 mil")',P),
('BYD Song Plus','2026','Zero-km','Híbrido plug-in (PHEV)',250000,'Tabela (aprox.)','2026 (mês não informado)','Motor Agora (cita "encosta nos R$ 250 mil")',P),
('Honda Civic Advanced Hybrid','2026','Zero-km','Híbrido (HEV)',266500,'Preço de venda','Out/2026','Despachante Dok',A),
('BYD Seal','2026','Zero-km','Elétrico (BEV)',299990,'Tabela sugerida','Out/2026','Despachante Dok (descontos de até R$ 30 mil em loja); FIPE zero-km jul/26: R$ 298.452',A)]
for i,(v,ano,ct,tc,pr,tp,ref,fo,vs) in enumerate(cat):
    r=5+i; put(c,r,1,f"V{i+1:02d}"); put(c,r,2,v,bold); put(c,r,3,ano); put(c,r,4,ct); put(c,r,5,tc); put(c,r,6,pr,inp,RS); put(c,r,7,tp); put(c,r,8,ref); put(c,r,9,fo); put(c,r,10,vs)
last=5+len(cat)-1
r=last+2
put(c,r,1,'Ind.'); put(c,r,2,'Preço médio de tabela dos carros novos (out/2026)',bold); put(c,r,3,''); put(c,r,4,'Mercado'); put(c,r,5,''); put(c,r,6,166142,inp,RS); put(c,r,7,'Média de tabela'); put(c,r,8,'Out/2026'); put(c,r,9,'Bright Consulting / Motor1 (arquivo recebido)'); put(c,r,10,B)
txt(c,r+2,'Não localizei preço para: SUV premium a gasolina, Volvo EX30 e BMW 330e (Dra. Helena), nem Tabela FIPE de elétrico ou híbrido usado de entrada (Ana Paula). O Palio, Celta, Corsa e Uno 2008 seguem a Tabela FIPE; o Celta é de 2004.')
txt(c,r+3,'Sites de FIPE usados reproduzem a Tabela FIPE oficial; o consulta oficial é veiculos.fipe.org.br. Preços de zero-km não incluem descontos, PcD nem frotistas.')
widths(c,[7,48,11,16,20,14,22,22,52,34]); c.freeze_panes='A5'
CAT_FIRST,CAT_LAST=5,last

# ============ P11_Personas ============
p=wb.create_sheet('P11_Personas')
title(p,'P11 – Personas: carros que cabem no orçamento (pergunta 3.3.2)','Cada opção usa o preço de P10 e os parâmetros de P9 (entrada, juros, prazo). Ajuste a entrada de cada linha para testar cenários.')
hdr(p,4,['Faixa','Persona','Tipo de opção','ID','Veículo','Categoria','Preço (R$)','Entrada (%)','Valor financiado (R$)','Parcela mensal (R$)','Renda mín. (R$)','Renda máx. (R$)','Parcela ÷ renda mín.','Parcela ÷ renda máx.','Teto do escopo – máx. (R$)','Preço ÷ teto do escopo','Situação','Comentário'],60)
opts=[(1,'Combustão / usado','V01','Opção mais barata do escopo (Uno Mille)'),(1,'Combustão / usado','V03','Palio 1.0 2008, citado no escopo'),(1,'Combustão / usado','V05','Corsa 2008, citado no escopo'),(1,'Eletrificada (contraste)','V23','Escopo: inviável; o preço confirma'),
(2,'Combustão / usado','V07','Onix 2019, versão de entrada'),(2,'Combustão / usado','V08','Onix 2019, versão LT'),(2,'Combustão / usado','V10','HB20S 2019, versão de entrada'),(2,'Eletrificada (contraste)','V23','Sem elétrico usado com preço FIPE; o zero-km foge do orçamento'),
(3,'Combustão 0 km','V21','Polo 0 km do escopo (versão Track)'),(3,'Seminovo','V12','Tracker 2022 seminovo do escopo'),(3,'Seminovo','V13','Tracker 2023 seminovo'),(3,'Eletrificada','V23','Dolphin Mini 0 km do escopo'),
(4,'Combustão 0 km','V27','T-Cross do escopo (versão Comfortline)'),(4,'Combustão 0 km','V29','Corolla Cross XRE em oferta'),(4,'Eletrificada seminovo','V16','Corolla Cross híbrido 2022 (alternativa)'),(4,'Eletrificada','V31','Haval H6 do escopo (valor aproximado)'),(4,'Eletrificada','V32','Song Plus do escopo (valor aproximado)'),
(5,'Eletrificada','V34','BYD Seal do escopo; EX30 e 330e sem preço localizado')]
# mapeamento de IDs: V23 = Dolphin Mini? checar pela lista
ids={nm:f"V{i+1:02d}" for i,(nm,*_) in enumerate(cat)}
fix={'V23':ids['BYD Dolphin Mini'],'V21':ids['Volkswagen Polo Track 1.0'],'V27':ids['Volkswagen T-Cross Comfortline'],'V29':ids['Toyota Corolla Cross XRE (estoque 2026, em oferta)'],
'V31':ids['GWM Haval H6 (a partir de)'],'V32':ids['BYD Song Plus'],'V34':ids['BYD Seal'],'V16':ids['Toyota Corolla Cross XRX 1.8 Híbrido']}
rows=[]
for fx,tp,idv,cm in opts:
    idv=fix.get(idv,idv); rows.append((fx,tp,idv,cm))
for i,(fx,tp,idv,cm) in enumerate(rows):
    r=5+i
    put(p,r,1,fx,fmt='0'); put(p,r,2,f"=INDEX(P9_Faixas_Renda!$B$28:$B$32,A{r})"); put(p,r,3,tp); put(p,r,4,idv,inp)
    rg=f"P10_Catalogo_Precos!$A${CAT_FIRST}:$A${CAT_LAST}"
    put(p,r,5,f"=INDEX(P10_Catalogo_Precos!$B${CAT_FIRST}:$B${CAT_LAST},MATCH(D{r},{rg},0))",link)
    put(p,r,6,f"=INDEX(P10_Catalogo_Precos!$D${CAT_FIRST}:$D${CAT_LAST},MATCH(D{r},{rg},0))",link)
    put(p,r,7,f"=INDEX(P10_Catalogo_Precos!$F${CAT_FIRST}:$F${CAT_LAST},MATCH(D{r},{rg},0))",link,RS)
    put(p,r,8,"=P9_Faixas_Renda!$B$9",link,'0%'); put(p,r,9,f"=G{r}*(1-H{r})",fmt=RS)
    put(p,r,10,f"=PMT(P9_Faixas_Renda!$B$7,P9_Faixas_Renda!$B$8,-I{r})",fmt=RS)
    put(p,r,11,f"=INDEX(P9_Faixas_Renda!$C$28:$C$32,A{r})",link,RS); put(p,r,12,f"=INDEX(P9_Faixas_Renda!$D$28:$D$32,A{r})",link,RS)
    put(p,r,13,f"=J{r}/K{r}",fmt=PC); put(p,r,14,f"=J{r}/L{r}",fmt=PC)
    put(p,r,15,f"=INDEX(P9_Faixas_Renda!$R$28:$R$32,A{r})",link,RS); put(p,r,16,f"=G{r}/O{r}",fmt='0.00"x"')
    put(p,r,17,f'=IF(G{r}>O{r},"Acima do teto do escopo",IF(M{r}<=P9_Faixas_Renda!$B$6,"Cabe na regra (mesmo na renda mínima)",IF(N{r}<=P9_Faixas_Renda!$B$6,"Cabe só na renda máxima da faixa","Excede a regra dos 20%")))',bold)
    put(p,r,18,cm); p.row_dimensions[r].height=30
PL=5+len(rows)-1
txt(p,PL+2,'Helena: renda e teto são valores mínimos ("acima de"), então "Cabe na regra" é conservador. Seu Zé: a Dolphin Mini aparece só para mostrar a distância do orçamento. O teto máximo do escopo para Ana Paula é R$ 65 mil.')
txt(p,PL+3,'Parcela = PMT(juros, prazo, financiado). O IPVA, seguro, combustível e manutenção não entram. Elétricos têm isenção ou desconto de IPVA em alguns estados (por exemplo, SP), o que não foi modelado.')
widths(p,[7,34,22,7,44,15,13,9,15,14,13,13,11,11,14,10,34,52]); p.freeze_panes='E5'

# ============ RESUMO: bloco 3.3 ============
rs=wb['RESUMO']
rs.cell(42,1,'3.3 Renda e personas').font=Font(name=F,size=11,bold=True,color=TEAL)
hdr(rs,43,['Indicador','Valor','Onde'],28)
res=[('Capacidade de compra – faixa 1 (Seu Zé), máx. pela regra','=P9_Faixas_Renda!P28',RS,'P9, bloco C'),
('Capacidade de compra – faixa 2 (Ana Paula), máx. pela regra','=P9_Faixas_Renda!P29',RS,'P9, bloco C'),
('Capacidade de compra – faixa 3 (Lucas & Mariana), máx.','=P9_Faixas_Renda!P30',RS,'P9, bloco C'),
('Capacidade de compra – faixa 4 (Roberto), máx.','=P9_Faixas_Renda!P31',RS,'P9, bloco C'),
('Entrada necessária para o teto de Ana Paula (R$ 65 mil)','=P9_Faixas_Renda!S29',PC,'P9, bloco C'),
('Entrada necessária para o teto de Lucas & Mariana (R$ 130 mil)','=P9_Faixas_Renda!S30',PC,'P9, bloco C'),
('% das famílias brasileiras com renda até 3 SM (POF 2017-18)','=P9_Faixas_Renda!J16',PC,'P9, bloco B'),
('Opções de persona dentro da regra, inclusive as que só cabem na renda máxima (nº de linhas)',f'=COUNTIF(P11_Personas!Q5:Q{PL},"Cabe*")','0','P11')]
for i,(a,f,fm,o) in enumerate(res):
    r=44+i; put(rs,r,1,a); put(rs,r,2,f,link,fm,teal_fill if i==0 else None); put(rs,r,3,o)
rs.cell(53,1,'• Pela regra dos 20%, os tetos do escopo de Ana Paula e de Lucas & Mariana só são alcançados com entrada alta; ver P9 e P11.').font=base
rs.cell(54,1,'• Faltam preços de elétrico usado (FIPE), EX30, BMW 330e e SUV premium; ver P10.').font=base

# ============ LEIA-ME ============
wl=wb['LEIA-ME']
wl['A1']='Panorama Automotivo – Etapas 2 e 3: Eletrificados, Renda e Personas (2021–2026)'
extra=[('',0),('ETAPA 3 – RENDA E PERSONAS (perguntas 3.3)',2),
('P9_Faixas_Renda – distribuição de renda das famílias (POF 2017-2018), regra dos 20%, capacidade de compra e faixa recomendada por nível de renda.',0),
('P10_Catalogo_Precos – preços de referência: Tabela FIPE (usados) e preços de tabela (zero-km), com fonte e situação da verificação.',0),
('P11_Personas – opções de carro por persona, parcela e situação frente à regra dos 20%.',0),
('Limites: POF de 2017-2018 (valores convertidos por salário mínimo); juros de financiamento de 2,04% a.m. são o mais recente que confirmei; IPVA, seguro, combustível e manutenção não entram; faltam alguns preços (ver P10).',0)]
r0=wl.max_row+1
for i,(t,k) in enumerate(extra):
    c=wl.cell(r0+i,1,t); c.alignment=Alignment(wrap_text=True,vertical='top'); c.font=Font(name=F,size=11 if k else 10,bold=bool(k),color=TEAL if k else '000000')

# ============ Auditoria: arquivo mercado_veiculos_eletrificados_brasil.xlsx ============
wa=wb['Auditoria_Fontes']
G='E2F4E3'; Y='FFF2CC'; R='F8D7DA'; Bl='DCE9F8'
st={'CONFERIDO':G,'RESSALVA':Y,'ERRO CORRIGIDO':R,'LACUNA PREENCHIDA':Bl}
n0=wa.max_row
items=[
('[Arquivo "mercado_veiculos_eletrificados_brasil"] Falta total para 3.3','LACUNA PREENCHIDA','O arquivo não tem dados da POF/IBGE, tabela FIPE de usados, nem as 5 personas. Só a aba Preços tem 8 valores de zero-km.','Criadas P9 (renda, regra dos 20%), P10 (preços e FIPE) e P11 (personas).'),
('Volume de 2022 e 2023 "estimado" e 2021 "n/d"','ERRO CORRIGIDO','O arquivo deriva 2023 e 2022 de taxas de crescimento e deixa 2021 sem dado. Os anuários FENABRAVE trazem os valores reais (automóveis): 2021 1.557.909; 2022 1.576.902; 2023 1.720.833.','Usar P1. O escopo do arquivo (automóveis + comerciais leves) é diferente do P1 (só automóveis); não misturar.'),
('Novos × usados: estimativa por 2 meses × 12','RESSALVA','Usados anuais = média de jan e mar/2025 × 12 (10,6 milhões), o que dá 4,2 usados por novo (80,6% usados). O Anuário 2025 diz 5,1 usados por novo (83,6%) para automóveis. Dois meses de uma série que mudou de metodologia em mar/2025 não sustentam o ano.','Usar P2 (razão publicada no anuário). Manter o arquivo apenas como comparação.'),
('Participação de eletrificados e tecnologia (2025, 1º sem/2026, set/2026)','CONFERIDO','Os números de BEV 80.178, PHEV 101.364, HEV 21.047, HEV Flex 21.323 e do 1º semestre batem com P4. 223.912 ÷ 2.549.462 = 8,8%, coerente com os 9% do arquivo.','Mantido.'),
('Ranking de modelos: Spark EUV 3.487 × 3.847 e Aion V 959 × 969','RESSALVA','O arquivo usa 3.487 e 959 (Canaltech/A Tarde), a planilha anterior usa 3.847 e 969 (Canal VE). 3.487 parece transposição de dígitos, mas não consegui abrir a fonte primária.','Mantidos os valores de P5 (Canal VE). Conferir na fonte.'),
('Haval H6: 18.508 no arquivo × 18.421 em P5','RESSALVA','A diferença de 87 unidades vem de agrupar as versões de formas distintas (A Tarde × soma das 4 linhas do Canal VE). Song Plus (13.762) e Corolla Cross (8.721) batem.','Sem alteração.'),
('Marcas: só 3 marcas de BEV em set/2026 (GAC 1.777 × 1.755)','RESSALVA','O arquivo diz que não há ranking oficial único; as abas P6 trazem o ranking completo de modelos, marcas e o anuário. GAC: 1.777 no arquivo e 1.755 na Federação.','Usar P6.'),
('Geografia do arquivo','CONFERIDO','Regiões, estados e cidades iguais a P7; 46,4% do Sudeste em 2025 confere. O 52% de 2023 não foi reverificado.','Mantido P7.'),
('Preços de zero-km (Mobi R$ 72.490 ... Dolphin Mini R$ 119.990)','RESSALVA','Os 8 valores vieram de reportagens (Bright Consulting, UOL, Auto+, iCarros) e não foram reverificados. Polo Track aparece como R$ 97.990 no arquivo, R$ 96.690 na Localiza e R$ 86.896 com desconto PcD/microempresa em fev/2026.','Mantidos em P10 com a marca "do arquivo recebido".'),
('Rótulo "Oficial" para dados de imprensa','RESSALVA','O arquivo marca como "Oficial" valores que foram lidos em reportagens que citam ABVE ou Fenabrave.','Em P10 a coluna "Situação da verificação" separa conferido, do arquivo e aproximado.'),
('Renda (POF) – sem tabela por classe de transporte','RESSALVA','O IBGE (SIDRA/FTP) bloqueou a leitura automática. Só obtive a distribuição das famílias por classe e a média nacional de transporte (cerca de 18%). A participação do transporte por classe de renda precisa ser baixada manualmente (tabelas 1.1.x de despesas da POF).','Usei a regra dos 20% do escopo; o gasto real por classe fica como próximo passo.'),
('Juros de financiamento','RESSALVA','A taxa mais recente que confirmei é 2,04% a.m. (27,4% a.a.), citada em jun/2026; não achei a de set/out 2026. Instituições variam de 0,6% a 4% a.m.','Parâmetro editável em P9 (célula azul).'),
('Tetos do escopo × regra dos 20%','ERRO CORRIGIDO','Com 48 meses, 2,04% a.m. e 20% de entrada, Ana Paula compra até R$ 42 mil (escopo: R$ 65 mil) e Lucas & Mariana até R$ 114 mil (escopo: R$ 130 mil). Os tetos só valem com entrada de cerca de 49% e 30%.','Mostrado em P9; ajustar o escopo ou deixar explícita a entrada.'),
('Nomes de personas e dados sem preço','LACUNA PREENCHIDA','Sem preço localizado: SUV premium a gasolina, Volvo EX30, BMW 330e e elétrico/híbrido usado de entrada.','Registrado em P10; precisa de consulta à Tabela FIPE oficial.')]
for i,(it,s,ev,ac) in enumerate(items):
    r=n0+1+i
    put(wa,r,1,20+i,fmt='0'); put(wa,r,2,it,bold); put(wa,r,3,s,bold,fill=PatternFill('solid',fgColor=st[s])); put(wa,r,4,ev); put(wa,r,5,ac)

# ============ Fontes: novas ============
wf=wb['Fontes']; r=wf.max_row+2
wf.cell(r,1,'N) Fontes acrescentadas na Etapa 3 (renda e personas)').font=Font(name=F,size=11,bold=True,color=TEAL); r+=1
new=[('N10','IBGE – POF 2017-2018: famílias por classe de rendimento','Pesquisa oficial','Classes de renda, 23,9% até 2 SM, 2,7% acima de 25 SM','https://agenciadenoticias.ibge.gov.br/agencia-noticias/25608-quase-25-da-renda-de-mais-pobres-vem-de-aposentadorias-e-programas-sociais.html','Lido via busca (o site bloqueou a leitura direta)'),
('N11','Carta Capital – POF: famílias por faixa de rendimento','Reportagem sobre a POF','12,8; 21,8; 9,6 e 4,4 milhões de famílias nas classes intermediárias','https://www.cartacapital.com.br/?p=95446','Lido via busca'),
('N12','IBGE – POF 2017-2018 (alimentação 17,5%, transporte)','Pesquisa oficial','Transporte ≈ 18,1% da despesa de consumo','https://agenciadenoticias.ibge.gov.br/agencia-sala-de-imprensa/2013-agencia-de-noticias/releases/25598-pof-2017-2018','Lido via busca'),
('N13','Decreto 12.797/2025 – salário mínimo 2026','Norma oficial (via imprensa)','R$ 1.621','https://exame.com/economia/salario-minimo-de-2026-quem-tem-direito-a-receber-o-reajuste/','Conferido por várias fontes'),
('N14','Conjur – juros de veículos (BCB): 2,04% a.m. e 27,42% a.a.','Reportagem que cita o BCB','Taxa de financiamento de veículos','https://www.conjur.com.br/2026-jun-23/bancos-devem-comprovar-risco-para-aplicar-juros-acima-da-media/','Lido via busca; série oficial: dadosabertos.bcb.gov.br (SGS 20749)'),
('N15','Tabela FIPE reproduzida por tabelafipebrasil.com e carzin.com.br','Reprodução da Tabela FIPE','Preços de usados (Onix, HB20, Palio, Uno, Corsa, Celta, Tracker, Corolla Cross)','https://www.tabelafipebrasil.com/carros/GM---CHEVROLET/modelo/ONIX/2019','Lido via busca; oficial: veiculos.fipe.org.br'),
('N16','Tribuna PR / Autoo / Canaltech – preços de zero-km','Reportagens','T-Cross, Tracker, Corolla Cross','https://www.tribunapr.com.br/noticias/automoveis/volkswagen-t-cross-novos-precos-reducao/','Lido via busca'),
('N17','Despachante Dok / Tabela Carros – BYD Seal e Civic Hybrid','Reportagens','Seal R$ 299.990; Civic Hybrid ≈ R$ 266.500','https://www.despachantedok.com.br/blog/?p=38738','Lido via busca'),
('N18','Motor Agora – BYD Song Plus 2026 × Haval H6','Reportagem','Preços aproximados (R$ 250 mil e R$ 239 mil)','https://www.motoragora.com.br/?p=186002','Lido via busca; valor aproximado')]
for row in new:
    vals=list(row)+['Consulta em 10/11/2026']
    vals=[vals[0],vals[1],vals[2],vals[3],vals[4],vals[5],'Consulta em 10/11/2026']
    for c_,v in enumerate(vals,1): put(wf,r,c_,v)
    r+=1

# ordem e cores
order=['LEIA-ME','RESUMO','P1_Volume_Anual','P1_Volume_Mensal','P2_Novos_vs_Usados','P3_Eletrificados','P4_Tecnologia','P5_Ranking_Modelos','P6_Ranking_Marcas','P7_Geografia','P8_Frota_Ref','P9_Faixas_Renda','P10_Catalogo_Precos','P11_Personas','Auditoria_Fontes','Fontes']
wb._sheets=[wb[n] for n in order]
for n in ('P9_Faixas_Renda','P10_Catalogo_Precos','P11_Personas'): wb[n].sheet_properties.tabColor='B8860B'
wb.save(OUT); print('saved',PL,CAT_LAST)
