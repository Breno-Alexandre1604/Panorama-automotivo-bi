import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

raw=json.load(open('/home/claude/monthly_raw.json'))
# edição mais recente de cada ano
monthly={2021:raw['23']['2021'],2022:raw['23']['2022'],2023:raw['25']['2023'],2024:raw['25']['2024'],2025:raw['25']['2025']}
edicao={2021:'Anuário FENABRAVE 2023',2022:'Anuário FENABRAVE 2023',2023:'Anuário FENABRAVE 2025',2024:'Anuário FENABRAVE 2025',2025:'Anuário FENABRAVE 2025'}
publicado={2021:1557909,2022:1576902,2023:1720833,2024:1948110,2025:1996519}
razao={2021:6.1,2022:5.5,2023:5.3,2024:5.1,2025:5.1}
razao_ed={2021:'Anuário 2021',2022:'Anuário 2022',2023:'Anuário 2023',2024:'Anuário 2024',2025:'Anuário 2025'}
bev={2021:2612,2022:7522,2023:18499,2024:60106,2025:79351}
hev={2021:32126,2022:40814,2023:74344,2024:114588,2025:203107}
elet_ed={2021:'Anuário 2023 (p.16)',2022:'Anuário 2023 (p.16)',2023:'Anuário 2025 (p.18)',2024:'Anuário 2025 (p.18)',2025:'Anuário 2025 (p.18)'}
meses=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
anos=[2021,2022,2023,2024,2025]

F='Arial'
hdr_fill=PatternFill('solid',fgColor='1B365D'); hdr_font=Font(name=F,bold=True,color='FFFFFF',size=10)
teal=PatternFill('solid',fgColor='D9EEEE')
base=Font(name=F,size=10); inp=Font(name=F,size=10,color='0000FF'); link=Font(name=F,size=10,color='008000'); bold=Font(name=F,size=10,bold=True)
thin=Side(style='thin',color='BFC5CF'); box=Border(top=thin,bottom=thin,left=thin,right=thin)
def header(ws,row,labels):
    for i,l in enumerate(labels,1):
        c=ws.cell(row,i,l); c.fill=hdr_fill; c.font=hdr_font; c.border=box
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[row].height=42
def widths(ws,w):
    for i,x in enumerate(w,1): ws.column_dimensions[get_column_letter(i)].width=x
def style_range(ws,r1,r2,c1,c2,font=base,fmt=None):
    for r in range(r1,r2+1):
        for c in range(c1,c2+1):
            cell=ws.cell(r,c); cell.font=font; cell.border=box
            if fmt: cell.number_format=fmt

wb=Workbook()
# ---------- LEIA-ME
ws=wb.active; ws.title='LEIA-ME'
lines=[
('Panorama Automotivo – Automóveis (2021–2025): Etapa 1 (FENABRAVE)',True),
('Disciplina: Análise de Dados e BI – UNIPÊ (Prof. Daniel Brandão) | Responsável por estas perguntas: Breno Alexandre',False),
('',False),
('ESCOPO',True),
('Somente o segmento AUTOMÓVEIS (emplacamentos de automóveis novos). Não inclui comerciais leves, caminhões, ônibus, motos nem implementos.',False),
('Perguntas respondidas: (1) Volume geral de vendas/emplacamentos por ano; (2) Proporção Novos vs. Usados; (3) Participação dos eletrificados vs. combustão.',False),
('',False),
('ABAS',True),
('P1_Volume_Anual – total de emplacamentos por ano (soma dos meses) e variação anual.',False),
('P1_Volume_Mensal – tabela fato mensal (formato longo, pronta para Power BI).',False),
('P2_Novos_vs_Usados – novos (emplacamentos), razão usados/novo, usados estimados e percentuais.',False),
('P3_Eletrificados – elétricos (BEV), híbridos, combustão e demais, e participações.',False),
('',False),
('FONTES (arquivos enviados)',True),
('2003_12_2.pdf = Anuário FENABRAVE 2021; Anuario2022.pdf; Anuario2023.pdf; Anuário2024.pdf; Anuario2025.pdf (dados de 2025; o título interno do PDF diz "anuario 2026").',False),
('Fonte primária dos dados citada nos anuários: SENATRAN. Para cada ano usou-se a edição mais recente do anuário em que ele aparece (coluna "Edição" nas abas).',False),
('',False),
('METODOLOGIA E LIMITAÇÕES (leia antes de usar nos gráficos)',True),
('1) Emplacamentos = automóveis novos (0 km) emplacados. É a medida de "vendas de novos". Os totais anuais aqui são a soma dos 12 meses e conferem com o "Acumulado" publicado.',False),
('2) Usados: o anuário NÃO publica o volume absoluto de usados no segmento. Publica a razão "usados negociados para cada novo emplacado". Usados estimados = Novos × Razão. A razão é arredondada em 1 casa decimal, então a estimativa tem margem de cerca de ±1%.',False),
('3) Cor das células: azul = dado digitado de um anuário; preto = fórmula; verde = vínculo com outra aba.',False),
('4) Eletrificados = elétricos (BEV) + híbridos, como o anuário reporta. O anuário não separa HEV, PHEV e híbrido leve (mild hybrid); essa abertura exigiria ABVE Data.',False),
('5) "Combustão e demais" = total de emplacamentos − eletrificados. Inclui flex, gasolina, diesel e outros. Não é uma série publicada, é derivada.',False),
('6) Inconsistências nas próprias fontes: o Anuário 2024 afirma que eletrificados são 5,4% do total, mas 174,7 mil ÷ 1,95 milhão dá ~9,0%. O gráfico de combustível do Anuário 2025 mostra 13,2% (texto: 14%) e o de 2024 mostra 10,30%, usando uma base diferente do total de emplacamentos. Aqui o percentual é sempre eletrificados ÷ total de emplacamentos de automóveis.',False),
('7) Pequenas revisões entre edições (poucas unidades) ocorrem porque a SENATRAN atualiza os dados. 2023 e 2024 de eletrificados diferem levemente entre os anuários 2023/2024 e 2025; foi usada a edição 2025.',False),
('8) Não foram usados ABVE Data nem dados.gov.br (não disponíveis para o grupo).',False),
]
for i,(t,b) in enumerate(lines,1):
    c=ws.cell(i,1,t); c.font=Font(name=F,size=12 if i==1 else 10,bold=b); c.alignment=Alignment(wrap_text=True,vertical='top')
ws.column_dimensions['A'].width=140

# ---------- P1_Volume_Mensal
wm=wb.create_sheet('P1_Volume_Mensal')
header(wm,1,['Ano','Mês (nº)','Mês','Emplacamentos de automóveis (un.)','Edição do anuário'])
r=2
for y in anos:
    for i,v in enumerate(monthly[y]):
        wm.cell(r,1,y); wm.cell(r,2,i+1); wm.cell(r,3,meses[i]); wm.cell(r,4,v); wm.cell(r,5,edicao[y]); r+=1
last_m=r-1
style_range(wm,2,last_m,1,5)
for rr in range(2,last_m+1):
    wm.cell(rr,4).font=inp; wm.cell(rr,4).number_format='#,##0'; wm.cell(rr,1).number_format='0'
widths(wm,[8,10,8,22,26]); wm.freeze_panes='A2'

# ---------- P1_Volume_Anual
wa=wb.create_sheet('P1_Volume_Anual',1)
header(wa,1,['Ano','Emplacamentos de automóveis (un.)','Total "Acumulado" publicado (un.)','Diferença (soma − publicado)','Variação % a.a.','Edição / observação'])
wa.cell(2,1,2020); wa.cell(2,2,1615528); wa.cell(2,6,'Referência para a variação de 2021 (Anuário 2021, Acumulado 2020)')
for i,y in enumerate(anos):
    r=3+i
    wa.cell(r,1,y)
    wa.cell(r,2,f"=SUMIFS(P1_Volume_Mensal!$D$2:$D${last_m},P1_Volume_Mensal!$A$2:$A${last_m},A{r})")
    wa.cell(r,3,publicado[y])
    wa.cell(r,4,f"=B{r}-C{r}")
    wa.cell(r,5,f"=B{r}/B{r-1}-1")
    wa.cell(r,6,edicao[y])
style_range(wa,2,7,1,6)
for r in range(2,8):
    wa.cell(r,1).number_format='0'
    for c in (2,3,4): wa.cell(r,c).number_format='#,##0;-#,##0;0'
    wa.cell(r,5).number_format='0.0%'
wa['B2'].font=inp
for r in range(3,8): wa.cell(r,3).font=inp
wa.cell(9,1,'Total 2021–2025'); wa.cell(9,2,'=SUM(B3:B7)')
wa.cell(10,1,'Média anual'); wa.cell(10,2,'=AVERAGE(B3:B7)')
wa.cell(11,1,'Crescimento 2021→2025'); wa.cell(11,2,'=B7/B3-1')
for r in (9,10,11):
    wa.cell(r,1).font=bold; wa.cell(r,2).font=bold; wa.cell(r,1).border=box; wa.cell(r,2).border=box
wa['B9'].number_format='#,##0'; wa['B10'].number_format='#,##0'; wa['B11'].number_format='0.0%'
wa.cell(13,1,'Fonte: FENABRAVE (Anuários 2021–2025), dados SENATRAN. Soma mensal conferida com o "Acumulado" publicado (coluna D deve ser 0).').font=Font(name=F,size=9,italic=True)
widths(wa,[22,24,24,20,16,58]); wa.freeze_panes='A2'

# ---------- P2
w2=wb.create_sheet('P2_Novos_vs_Usados')
header(w2,1,['Ano','Novos 0 km – emplacamentos (un.)','Razão usados por novo','Usados estimados (un.)','Total negociado (un.)','% Novos','% Usados','Fonte da razão'])
for i,y in enumerate(anos):
    r=2+i
    w2.cell(r,1,y)
    w2.cell(r,2,f"=P1_Volume_Anual!B{3+i}")
    w2.cell(r,3,razao[y])
    w2.cell(r,4,f"=B{r}*C{r}")
    w2.cell(r,5,f"=B{r}+D{r}")
    w2.cell(r,6,f"=B{r}/E{r}")
    w2.cell(r,7,f"=D{r}/E{r}")
    w2.cell(r,8,razao_ed[y]+' – proporção nacional (Brasil)')
style_range(w2,2,6,1,8)
for r in range(2,7):
    w2.cell(r,1).number_format='0'; w2.cell(r,2).font=link; w2.cell(r,3).font=inp
    w2.cell(r,2).number_format='#,##0'; w2.cell(r,3).number_format='0.0'
    w2.cell(r,4).number_format='#,##0'; w2.cell(r,5).number_format='#,##0'
    w2.cell(r,6).number_format='0.0%'; w2.cell(r,7).number_format='0.0%'
w2.cell(8,1,'Acumulado 2021–2025').font=bold
w2.cell(8,2,'=SUM(B2:B6)'); w2.cell(8,4,'=SUM(D2:D6)'); w2.cell(8,5,'=SUM(E2:E6)'); w2.cell(8,6,'=B8/E8'); w2.cell(8,7,'=D8/E8')
for c in range(1,8): w2.cell(8,c).border=box; 
for c in (2,4,5): w2.cell(8,c).number_format='#,##0'; w2.cell(8,c).font=bold
for c in (6,7): w2.cell(8,c).number_format='0.0%'; w2.cell(8,c).font=bold
w2.cell(10,1,'Nota: Usados estimados = Novos × Razão. O anuário publica só a razão ("usados negociados para cada novo emplacado"), arredondada em 1 casa; margem de ~±1%.').font=Font(name=F,size=9,italic=True)
widths(w2,[20,26,16,22,22,11,11,44]); w2.freeze_panes='A2'

# ---------- P3
w3=wb.create_sheet('P3_Eletrificados')
header(w3,1,['Ano','Elétricos BEV (un.)','Híbridos (un.)','Total eletrificados (un.)','Total emplacamentos de automóveis (un.)','Combustão e demais (un.)','% Elétricos','% Híbridos','% Eletrificados','% Combustão e demais','Var. % eletrificados a.a.','Fonte (eletrificados)'])
for i,y in enumerate(anos):
    r=2+i
    w3.cell(r,1,y); w3.cell(r,2,bev[y]); w3.cell(r,3,hev[y])
    w3.cell(r,4,f"=B{r}+C{r}")
    w3.cell(r,5,f"=P1_Volume_Anual!B{3+i}")
    w3.cell(r,6,f"=E{r}-D{r}")
    w3.cell(r,7,f"=B{r}/E{r}"); w3.cell(r,8,f"=C{r}/E{r}"); w3.cell(r,9,f"=D{r}/E{r}"); w3.cell(r,10,f"=F{r}/E{r}")
    w3.cell(r,11,'' if i==0 else f"=D{r}/D{r-1}-1")
    w3.cell(r,12,elet_ed[y])
style_range(w3,2,6,1,12)
for r in range(2,7):
    w3.cell(r,1).number_format='0'
    w3.cell(r,2).font=inp; w3.cell(r,3).font=inp; w3.cell(r,5).font=link
    for c in (2,3,4,5,6): w3.cell(r,c).number_format='#,##0'
    for c in (7,8,9,10,11): w3.cell(r,c).number_format='0.0%'
    w3.cell(r,9).fill=teal
w3.cell(8,1,'Nota: % sobre o total de emplacamentos de automóveis (SENATRAN). Os gráficos de combustível dos anuários usam outra base (ex.: 13,2% em 2025, 10,30% em 2024) e o Anuário 2024 cita 5,4%; ver aba LEIA-ME, item 6.').font=Font(name=F,size=9,italic=True)
w3.cell(9,1,'2021 e 2022: Anuário 2023. 2023–2025: Anuário 2025 (edição mais recente, com pequenas revisões).').font=Font(name=F,size=9,italic=True)
widths(w3,[8,16,16,18,24,20,11,11,13,15,16,24]); w3.freeze_panes='A2'
wb.save('/mnt/user-data/outputs/FENABRAVE_Automoveis_2021-2025_Etapa1.xlsx')
print('ok',last_m)
