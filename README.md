# 🛒 Comparador de Preus Supermercats

Script Python que cada mes llegeix una llista de la compra, busca preus a 4 supermercats online i genera un Excel comparatiu indicant on comprar cada producte per estalviar al màxim.

## Supermercats

| Supermercat | Ports | Ports gratis a partir de |
|-------------|-------|--------------------------|
| Mercadona | 7.90€ | — |
| BonPreu | 4.90€ | — |
| Alcampo | 4.90€ | 49€ |
| Caprabo | 4.90€ | 60€ |

## Requisits

```bash
pip install playwright openpyxl
playwright install chromium
```

## Ús

1. Edita `llista.txt` amb els productes que vols comprar
2. Executa l'script:

```bash
python comparador_preus_v7.py
```

O fes doble clic a `executar.bat` (Windows).

L'script genera un Excel `comparativa_preus_YYYYMMDD.xlsx` amb 3 fulls:

- **Comparativa** — tots els productes amb preus per supermercat
- **Per Supermercat** — agrupat per on has d'anar a comprar
- **Resum Estalvi** — comparativa de costos totals

## Format llista.txt

```
# Les línies amb # s'ignoren
MARCA Nom del producte xQUANTITAT

# Exemples:
BONPREU Arròs extra x5
HELLMANN'S Maionesa x1
DAURA Cervesa apta per celíacs en llauna x20
```

## Temps d'execució

~10-20 minuts per 48 productes (fa pauses entre cerques per no ser blocat).