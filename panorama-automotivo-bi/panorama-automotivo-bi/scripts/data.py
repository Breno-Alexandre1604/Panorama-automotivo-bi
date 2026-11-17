# Dados do Luis (conferidos) + complementos verificados
BEV,PHEV,FLEX,HEV='Elétrico (BEV)','Híbrido plug-in (PHEV)','Híbrido flex (HEV Flex)','Híbrido convencional (HEV)'
SRC_RANK='ABVE via Canal VE (08/07/2026)'
models=[ # tech, pos, modelo, marca, familia, unidades
(BEV,1,'Dolphin Mini','BYD','Dolphin Mini',35680),(BEV,2,'Dolphin','BYD','Dolphin',18054),(BEV,3,'EX2','Geely','EX2',14780),
(BEV,4,'Spark EUV','Chevrolet','Spark EUV',3847),(BEV,5,'Yuan Pro','BYD','Yuan Pro',2668),(BEV,6,'Ora 03','GWM','Ora 03',1837),
(BEV,7,'EX30','Volvo','EX30',1346),(BEV,8,'EX5','Geely','EX5',1343),(BEV,9,'Aion V','GAC','Aion V',969),(BEV,10,'Captiva EV','Chevrolet','Captiva EV',928),
(PHEV,1,'Song Pro','BYD','Song Pro',17776),(PHEV,2,'Song Plus','BYD','Song Plus',10951),(PHEV,3,'King','BYD','King',8205),
(PHEV,4,'Haval H6 híbrido plug-in (versão 19)','GWM','Haval H6',6918),(PHEV,5,'Jaecoo 7','Jaecoo','Jaecoo 7',6425),(PHEV,6,'Tank 300','GWM','Tank 300',2869),
(PHEV,7,'Song Plus Premium','BYD','Song Plus',2811),(PHEV,8,'C10 com autonomia estendida','Leapmotor','C10',2465),
(PHEV,9,'Haval H6 híbrido plug-in (versão 35)','GWM','Haval H6',2175),(PHEV,10,'EX5 híbrido plug-in','Geely','EX5',2020),
(FLEX,1,'Corolla Cross XRX Hybrid','Toyota','Corolla Cross',8721),(FLEX,2,'Corolla','Toyota','Corolla',7886),(FLEX,3,'Yaris Cross XRX Hybrid','Toyota','Yaris Cross',4686),
(FLEX,4,'Yaris Cross XRE Hybrid','Toyota','Yaris Cross',2384),(FLEX,5,'Haval H6 híbrido flex','GWM','Haval H6',401),
(HEV,1,'Omoda 5','Omoda','Omoda 5',9184),(HEV,2,'Haval H6 híbrido','GWM','Haval H6',8927),(HEV,3,'GS4','GAC','GS4',2130),(HEV,4,'RAV4','Toyota','RAV4',1286),
(HEV,5,'Maverick','Ford','Maverick',799),(HEV,6,'Kona','Hyundai','Kona',373),(HEV,7,'NX 350h','Lexus','NX 350h',326),(HEV,8,'Koleos','Renault','Koleos',188),
(HEV,9,'Niro','Kia','Niro',179),(HEV,10,'UX 300h','Lexus','UX 300h',127)]
# anuário FENABRAVE 2025 p.18 (conferido por coordenadas)
brand_bev=[('BYD',10217,43754,57084),('VOLVO',2643,4338,5162),('GEELY',None,1,3370),('GWM',775,6324,3237),('GM',203,227,2210),('RENAULT',435,1242,1728),
('BMW',1132,1024,944),('JAC',693,1333,540),('MINI',396,306,533),('PORSCHE',322,319,468),('AUDI',304,151,224),('M.BENZ',156,148,178),('NISSAN',370,63,101),('FORD',123,138,63),('PEUGEOT',269,354,4)]
brand_hyb=[('BYD',7664,32594,54577),('FIAT',None,3827,42601),('GWM',10702,22892,36062),('TOYOTA',21038,20353,24210),('M.BENZ',1370,6199,8705),('CAOA CHERY',11684,7258,6199),
('VOLVO',5535,4290,4556),('BMW',3303,3694,3498),('LAND ROVER',3718,3065,2276),('HONDA',577,3155,2149),('PORSCHE',1722,1566,2114),('PEUGEOT',None,None,1866),
('LEXUS',850,1100,1297),('MITSUBISHI',2,None,1086),('KIA',2383,2286,1028)]
