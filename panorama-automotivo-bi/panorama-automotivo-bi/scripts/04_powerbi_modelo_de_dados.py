import sys, datetime as dt
sys.path.insert(0,'/home/claude/b2')
from data import *
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill
src=load_workbook('/mnt/user-data/outputs/Projeto_BI_Automotivo_Etapa3_Renda_Personas.xlsx',data_only=True)
T={}
# --- fato_emplacamentos_mensal
m=src['P1_Volume_Mensal']; rows=[]
for r in range(2,62):
    a,mn,ms,v,ed=[m.cell(r,c).value for c in range(1,6)]
    rows.append([dt.date(int(a),int(mn),1),int(a),int(mn),ms,int(v)])
T['fato_emplacamentos_mensal']=(['Data','Ano','MesNum','Mes','Emplacamentos'],rows)
# --- dim_calendario
T['dim_calendario']=(['Data','Ano','MesNum','Mes','Trimestre'],[[r[0],r[1],r[2],r[3],f"T{(r[2]-1)//3+1}"] for r in rows])
T['dim_ano']=(['Ano'],[[y] for y in range(2021,2026)])
# --- fato_mercado_anual
p1=src['P1_Volume_Anual']; p2=src['P2_Novos_vs_Usados']; p3=src['P3_Eletrificados']; rows=[]
for i in range(5):
    y=2021+i
    nov=p1.cell(3+i,2).value; var=p1.cell(3+i,5).value
    raz=p2.cell(2+i,3).value; usd=p2.cell(2+i,4).value
    bev=p3.cell(2+i,2).value; hib=p3.cell(2+i,3).value
    rows.append([y,nov,var,raz,usd,nov+usd,nov/(nov+usd),usd/(nov+usd),bev,hib,bev+hib,nov-(bev+hib),(bev+hib)/nov])
T['fato_mercado_anual']=(['Ano','Emplacamentos_Automoveis','Var_Anual','Razao_Usados_por_Novo','Usados_Estimados','Total_Negociado','Pct_Novos','Pct_Usados','BEV','Hibridos_com_MHEV','Eletrificados_Total','Combustao_e_Demais','Pct_Eletrificados'],rows)
# --- dim_tecnologia
T['dim_tecnologia']=(['Tecnologia','Grupo','Criterio','Ordem'],[['BEV','Elétrico','ABVE / anuário',1],['PHEV','Híbrido','ABVE',2],['HEV','Híbrido','ABVE',3],['HEV Flex','Híbrido','ABVE',4],['Híbrido (c/ MHEV)','Híbrido','Anuário FENABRAVE (inclui micro-híbridos)',5]])
# --- fato_eletrificados_abve
p4=src['P4_Tecnologia']; rows=[]
per=[(15,'2024 (c/ MHEV no total divulgado)',1),(16,'2025',2),(17,'1º sem/2026',3),(18,'Junho/2026',4),(19,'Setembro/2026',5)]
for r,nm,o in per:
    for c,t in zip((2,3,4,5),('BEV','PHEV','HEV','HEV Flex')): rows.append([nm,o,t,p4.cell(r,c).value])
T['fato_eletrificados_abve']=(['Periodo','OrdemPeriodo','Tecnologia','Unidades'],rows)
# --- fato_modelos
p5=src['P5_Ranking_Modelos']; tmap={BEV:'BEV',PHEV:'PHEV',FLEX:'HEV Flex',HEV:'HEV'}; rows=[]
for r in range(5,40):
    t,pos,mod,mar,fam,u=[p5.cell(r,c).value for c in (2,1,3,4,5,6)]
    rows.append([tmap[t],pos,mod,mar,fam,u,'1º sem/2026'])
T['fato_modelos']=(['Tecnologia','PosicaoNaTecnologia','Modelo','Marca','Familia','Unidades','Periodo'],rows)
# --- fato_marcas_anuario
rows=[]
for tec,lst in (('BEV',brand_bev),('Híbrido (c/ MHEV)',brand_hyb)):
    for mar,a,b,c in lst:
        for ano,v in zip((2023,2024,2025),(a,b,c)): rows.append([mar,tec,ano,0 if v is None else v])
T['fato_marcas_anuario']=(['Marca','Tecnologia','Ano','Unidades'],rows)
# --- dim_regiao_uf
uf=[('AC','Acre','Norte'),('AL','Alagoas','Nordeste'),('AP','Amapá','Norte'),('AM','Amazonas','Norte'),('BA','Bahia','Nordeste'),('CE','Ceará','Nordeste'),('DF','Distrito Federal','Centro-Oeste'),('ES','Espírito Santo','Sudeste'),('GO','Goiás','Centro-Oeste'),('MA','Maranhão','Nordeste'),('MT','Mato Grosso','Centro-Oeste'),('MS','Mato Grosso do Sul','Centro-Oeste'),('MG','Minas Gerais','Sudeste'),('PA','Pará','Norte'),('PB','Paraíba','Nordeste'),('PR','Paraná','Sul'),('PE','Pernambuco','Nordeste'),('PI','Piauí','Nordeste'),('RJ','Rio de Janeiro','Sudeste'),('RN','Rio Grande do Norte','Nordeste'),('RS','Rio Grande do Sul','Sul'),('RO','Rondônia','Norte'),('RR','Roraima','Norte'),('SC','Santa Catarina','Sul'),('SP','São Paulo','Sudeste'),('SE','Sergipe','Nordeste'),('TO','Tocantins','Norte')]
T['dim_regiao_uf']=(['UF','Estado','Regiao'],[list(x) for x in uf])
# --- fato_geografia
g=src['P7_Geografia']; rows=[]
for r in range(5,10):
    rows.append(['Região',g.cell(r,1).value,None,g.cell(r,1).value,'1º sem/2026',g.cell(r,2).value]); rows.append(['Região',g.cell(r,1).value,None,g.cell(r,1).value,'2025',g.cell(r,4).value])
for r in range(15,20):
    n,u,rg=[g.cell(r,c).value for c in (1,2,3)]
    rows.append(['Estado',n,u,rg,'1º sem/2026',g.cell(r,4).value]); rows.append(['Estado',n,u,rg,'2025',g.cell(r,6).value])
for r in range(25,30):
    rows.append(['Cidade',g.cell(r,1).value,g.cell(r,2).value,g.cell(r,3).value,'1º sem/2026',g.cell(r,4).value])
T['fato_geografia']=(['Nivel','Local','UF','Regiao','Periodo','Emplacamentos'],rows)
# --- dim_classe_pof
p9=src['P9_Faixas_Renda']; rows=[]
for r in range(15,22):
    lo=p9.cell(r,2).value; hi=p9.cell(r,3).value
    rows.append([r-14,p9.cell(r,1).value,lo,hi if isinstance(hi,(int,float)) else None,hi if isinstance(hi,(int,float)) else 999,p9.cell(r,8).value,p9.cell(r,9).value,p9.cell(r,10).value])
T['dim_classe_pof']=(['ClasseNum','Classe','SM_inferior','SM_superior','SM_superior_calc','Familias_milhoes','Pct_Familias','Pct_Acumulado'],rows)
# --- dim_faixa_renda
w={1:('Seu Zé / Cláudio','Mobilidade essencial','Fiat Palio (2008), Chevrolet Celta, Corsa, Uno Mille','Inviável (foco em transporte público / custo mínimo)'),
2:('Ana Paula','Entrada / seminovo','Chevrolet Onix / HB20 usado (2018-2021)','Usado de entrada (em transição)'),
3:('Lucas & Mariana','Classe média','VW Polo 0 km / Tracker seminovo','BYD Dolphin Mini 0 km'),
4:('Roberto','Executivo','VW T-Cross / Corolla Cross','BYD Song Plus / Haval H6 (PHEV)'),
5:('Dra. Helena','Alta renda','SUV premium a gasolina','BYD Seal / Volvo EX30 / BMW 330e')}
rows=[]
for i in range(5):
    r=28+i; n,perf,comb,ele=w[i+1]
    rows.append([i+1,n,perf,p9.cell(r,3).value,p9.cell(r,4).value,p9.cell(r,5).value,p9.cell(r,17).value,p9.cell(r,18).value,p9.cell(r,15).value,p9.cell(r,16).value,p9.cell(r,8).value,p9.cell(r,9).value,comb,ele])
T['dim_faixa_renda']=(['FaixaNum','Persona','Perfil','Renda_min','Renda_max','Renda_media','Teto_escopo_min','Teto_escopo_max','Preco_max_regra_min','Preco_max_regra_max','Classe_POF','Pct_Familias_Classe','Opcao_Combustao_Usado','Opcao_Eletrificada'],rows)
# --- dim_veiculo
p10=src['P10_Catalogo_Precos']; rows=[]
for r in range(5,39):
    idv,v,ano,ct,tc,pr,tp,ref,fo,vs=[p10.cell(r,c).value for c in range(1,11)]
    prop='Elétrico' if 'BEV' in tc else ('Híbrido' if 'brido' in tc else 'Combustão')
    novo='Zero-km' if ct=='Zero-km' else ('Usado/seminovo')
    rows.append([idv,f"{v} ({ano})",v,ano,ct,novo,tc,prop,pr,tp,ref,fo,vs])
T['dim_veiculo']=(['ID','Veiculo_Ano','Veiculo','AnoModelo','Categoria','Condicao','Tecnologia','Propulsao','Preco','TipoPreco','Referencia','Fonte','Verificacao'],rows)
# --- escreve
wb=Workbook(); wb.remove(wb.active)
for name,(cols,data) in T.items():
    ws=wb.create_sheet(name[:31]); ws.append(cols)
    for row in data: ws.append(row)
    for c in ws[1]: c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='1B365D')
    ref=f"A1:{chr(64+len(cols)) if len(cols)<27 else 'Z'}{len(data)+1}"
    t=Table(displayName=name,ref=ref); t.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True); ws.add_table(t)
    for i,c in enumerate(cols,1): ws.column_dimensions[chr(64+i)].width=max(12,min(40,len(c)+4))
    if 'Data' in cols:
        for r in range(2,len(data)+2): ws.cell(r,1).number_format='yyyy-mm-dd'
wb.save('/mnt/user-data/outputs/PowerBI_Modelo_de_Dados.xlsx')
print({k:len(v[1]) for k,v in T.items()})
# --- verificação em python da lógica DAX (exemplos)
cat=T['dim_veiculo'][1]
def sim(renda,ent,n,i=0.0204,regra=0.2):
    f=(1-(1+i)**-n)/i; pm=renda*regra; pmax=pm*f/(1-ent)
    ok=[(r[8],r[1],r[8]*(1-ent)/f) for r in cat if r[8]*(1-ent)/f<=pm]
    best=max(ok) if ok else None
    return round(pm),round(pmax),len(ok),best
print(sim(5000,0.2,48)); print(sim(12000,0.3,60)); print(sim(25000,0.2,48)); print(sim(3000,0.2,48))
