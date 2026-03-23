"""
Script para generar datos de ejemplo en Excel.
Ejecutar una sola vez: python crear_datos_ejemplo.py
"""
import pandas as pd
import random
from datetime import date, timedelta

random.seed(42)

VENDEDORES = ["Carlos Medina", "Lucia Torres", "Jorge Quispe", "Ana Flores"]
DISTRITOS  = ["Miraflores", "San Isidro", "Surco", "La Molina", "Barranco", "Chorrillos"]

# ---------- clientes (MANTENIMIENTO) ----------
CLIENTES = [
    "DISTRIBUIDORA NORTE SAC", "IMPORTACIONES SUR EIRL", "COMERCIAL CENTRO SA",
    "BODEGA LOS ANDES", "FERRETERÍA EL SOL", "SUPERMERCADO PACÍFICO",
    "DISTRIBUCIONES LIMA SAC", "GRUPO COMERCIAL ANDINO",
]
MOT_MANT   = ["TOMAR PEDIDO", "CAPACITACIÓN", "LANZAMIENTO", "COBRANZA", "RECLAMO", "OTROS"]
MOT_MANT_N = [None] * len(MOT_MANT)   # MOTIVO NRO no aplica para Mantenimiento

# ---------- prospectos (PROSPECCIÓN) ----------
PROSPECTOS = [
    "MEGA MARKET PERU SAC", "INVERSIONES DEL NORTE EIRL", "COMERCIO RAPIDO SRL",
    "TIENDA NUEVA SAC", "DISTRIBUIDORA CENTRAL EIRL", "PUNTO VENTA SAC",
]
MOT_PROS   = ["PROSPECCIÓN", "CALIFICACIÓN DE LEADS", "VISITA", "PROPUESTA", "NEGOCIACIÓN", "CIERRE", "NO CIERRE"]
MOT_PROS_N = [1, 2, 3, 4, 5, 6, 7]

MESES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio"}

filas = []
inicio = date(2024, 1, 8)

for semana in range(24):          # 24 semanas = ~6 meses
    for vendedor in VENDEDORES:
        # --- Mantenimiento (4-7 visitas por semana) ---
        for _ in range(random.randint(4, 7)):
            cliente = random.choice(CLIENTES)
            idx     = random.randint(0, len(MOT_MANT)-1)
            motivo  = MOT_MANT[idx]
            delta   = timedelta(days=random.randint(0, 4))
            d       = inicio + timedelta(weeks=semana) + delta
            filas.append({
                "VENDEDOR":            vendedor,
                "FECHA":               d,
                "MES":                 MESES[d.month] if d.month in MESES else "Julio",
                "TIPO DE VISITA":      "MANTENIMIENTO",
                "TIPO DE CLIENTE":     "CLIENTE",
                "RAZON SOCIAL CLIENTE":cliente,
                "DISTRITO":            random.choice(DISTRITOS),
                "CONTACTO":            "Contacto " + cliente[:6],
                "TELÉFONO":            f"9{random.randint(10000000,99999999)}",
                "MOTIVO VISITA":       motivo,
                "MOTIVO NRO":          "",
                "RESULTADO / OBS":     "OK",
            })

        # --- Prospección (2-4 visitas por semana) ---
        for _ in range(random.randint(2, 4)):
            prospecto = random.choice(PROSPECTOS)
            idx       = random.randint(0, len(MOT_PROS)-1)
            motivo    = MOT_PROS[idx]
            nro       = MOT_PROS_N[idx]
            delta     = timedelta(days=random.randint(0, 4))
            d         = inicio + timedelta(weeks=semana) + delta
            filas.append({
                "VENDEDOR":            vendedor,
                "FECHA":               d,
                "MES":                 MESES[d.month] if d.month in MESES else "Julio",
                "TIPO DE VISITA":      "PROSPECCIÓN",
                "TIPO DE CLIENTE":     "PROSPECTO",
                "RAZON SOCIAL CLIENTE":prospecto,
                "DISTRITO":            random.choice(DISTRITOS),
                "CONTACTO":            "Lead " + prospecto[:6],
                "TELÉFONO":            f"9{random.randint(10000000,99999999)}",
                "MOTIVO VISITA":       motivo,
                "MOTIVO NRO":          nro,
                "RESULTADO / OBS":     random.choice(["Interesado", "Pendiente seguimiento", "No disponible"]),
            })

df = pd.DataFrame(filas).sort_values("FECHA").reset_index(drop=True)
df.to_excel("visitas_ventas.xlsx", index=False, sheet_name="Visitas")
print(f"✅  Archivo 'visitas_ventas.xlsx' creado con {len(df)} registros.")
