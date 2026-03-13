╔══════════════════════════════════════════════════════════════╗
║         COMPARADOR DE PREUS SUPERMERCATS  v7.0               ║
║         (C) 2026  -  Masnou, Catalunya                       ║
╚══════════════════════════════════════════════════════════════╝

  ■ DESCRIPCIO DEL PROGRAMA
  ─────────────────────────
  Llegeix una llista de la compra (LLISTA.TXT), consulta preus
  a 4 supermercats online i genera un Excel comparatiu amb
  el supermercat mes barat per cada producte.


  ■ SUPERMERCATS SUPORTATS
  ────────────────────────
  > MERCADONA   ports: 7.90€   ports gratis: NO
  > BONPREU     ports: 4.90€   ports gratis: NO
  > ALCAMPO     ports: 4.90€   ports gratis: +49.00€
  > CAPRABO     ports: 4.90€   ports gratis: +60.00€


  ■ REQUISITS DEL SISTEMA
  ───────────────────────
  - Python 3.10 o superior
  - Connexio a Internet


  ■ INSTALLACIO
  ─────────────
  C:\> pip install -r requirements.txt
  C:\> playwright install chromium


  ■ EXECUCIO
  ──────────
  C:\> python comparador_preus_v7.py

       -- o be --

  Fes doble clic a EXECUTAR.BAT


  ■ FORMAT DE LLISTA.TXT
  ──────────────────────
  # Les linies amb # s'ignoren
  # Format: MARCA Nom del producte xQUANTITAT

  BONPREU Arros extra x5
  HELLMANN'S Maionesa x1
  DAURA Cervesa apta per celiacs en llauna x20


  ■ FITXERS GENERATS
  ──────────────────
  comparativa_preus_YYYYMMDD.xlsx

    Full 1 │ COMPARATIVA    tots els productes i preus
    Full 2 │ PER SUPERMERCAT agrupat per on comprar
    Full 3 │ RESUM ESTALVI  comparativa costos totals


  ■ TEMPS D'EXECUCIO
  ──────────────────
  ~10-20 minuts per 48 productes
  (pauses entre cerques per evitar bloqueig)


  ■ AVIS
  ──────
  Aquest programa utilitza Playwright per navegar pels webs
  dels supermercats. Els preus poden variar en qualsevol
  moment. Verifica sempre abans de fer la comanda.


══════════════════════════════════════════════════════════════
  Prem qualsevol tecla per continuar . . . _
══════════════════════════════════════════════════════════════
