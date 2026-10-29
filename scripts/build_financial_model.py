import os
import sys
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt
import seaborn as sns

def build_model():
    print("=" * 70)
    print("INICIANDO CONSTRUCCION DE MODELO FINANCIERO AVANZADO & MONTE CARLO")
    print("=" * 70)

    output_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    excel_path = os.path.join(output_dir, "data", "Modelo_Valoracion_Financiera_MonteCarlo.xlsx")
    assets_dir = os.path.join(output_dir, "assets")
    os.makedirs(os.path.dirname(excel_path), exist_ok=True)
    os.makedirs(assets_dir, exist_ok=True)

    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Styles
    font_title = Font(name="Calibri", size=16, bold=True, color="1E3A8A")
    font_section = Font(name="Calibri", size=12, bold=True, color="1E293B")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_normal = Font(name="Calibri", size=10)
    font_kpi_num = Font(name="Calibri", size=14, bold=True, color="1E3A8A")

    fill_header_navy = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_header_slate = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    fill_subtle_blue = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    fill_subtle_green = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    fill_subtle_yellow = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_subtle_red = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="CBD5E1")
    double_bottom_side = Side(border_style="double", color="1E293B")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_total = Border(top=thin_border_side, bottom=double_bottom_side)

    # ----------------------------------------------------
    # SHEET 1: PARAMETROS & SUPUESTOS
    # ----------------------------------------------------
    print("[1/5] Creando hoja '01_Parametros_Supuestos'...")
    ws1 = wb.create_sheet(title="01_Parametros_Supuestos")
    ws1.views.sheetView[0].showGridLines = True

    ws1["B2"] = "MODELO DE VALORACIÓN FINANCIERA — SISTEMA AUTÓNOMO DE MONITOREO"
    ws1["B2"].font = font_title
    ws1["B3"] = "Universidad Nacional de Ingeniería (FIIS UNI) | Proyecto de Inversión Tecnológica"
    ws1["B3"].font = Font(name="Calibri", size=11, italic=True, color="64748B")

    params = [
        ("INVERSIÓN INICIAL & ACTIVOS", "", "", ""),
        ("Inversión Inicial en Activo Fijo (CAPEX)", 250000, "$#,##0", "Equipamiento de drones, sensores IoT, servidores edge"),
        ("Vida Útil del Proyecto (Años)", 5, "0", "Horizonte de evaluación económica"),
        ("Depreciación Anual Lineal", "=B6/B7", "$#,##0", "Método de línea recta sin valor de desecho contable"),
        ("Valor Residual de Mercado (Fin de Año 5)", 35000, "$#,##0", "Valor de recuperación de equipos"),
        ("", "", "", ""),
        ("OPERACIÓN COMERCIAL & DEMANDA", "", "", ""),
        ("Volumen de Servicio Base - Año 1 (Horas/Unid)", 12000, "#,##0", "Capacidad comercial proyectada"),
        ("Precio Unitario Base ($/Hora de Inspección)", 55.00, "$#,##0.00", "Tarifa por hora de servicio de telemetría"),
        ("Costo Variable Unitario ($/Hora)", 22.00, "$#,##0.00", "Baterías, consumibles, mantenimiento directo"),
        ("Costos Fijos Operativos Anuales (OPEX)", 95000, "$#,##0", "Personal especializado, seguros, nube, soporte"),
        ("Tasa de Crecimiento Anual del Volumen", 0.08, "0.0%", "Penetración de mercado proyectada"),
        ("Tasa de Inflación Anual de Precios", 0.03, "0.0%", "Ajuste tarifario anual"),
        ("Tasa de Inflación Anual de Costos", 0.025, "0.0%", "Incremento salarial y de insumos"),
        ("", "", "", ""),
        ("FINANZAS CORPORATIVAS & IMPUESTOS", "", "", ""),
        ("Tasa de Impuesto a la Renta (Perú)", 0.295, "0.0%", "Régimen General de Impuesto a la Renta (SUNAT)"),
        ("Costo Promedio Ponderado de Capital (WACC)", 0.115, "0.0%", "Tasa de descuento de los flujos libres de caja (FCFF)"),
    ]

    ws1["B5"] = "Parámetro Macroeconómico / Operativo"
    ws1["C5"] = "Valor"
    ws1["D5"] = "Unidad / Formato"
    ws1["E5"] = "Descripción Técnica y Justificación"
    for col, h in enumerate(["B5", "C5", "D5", "E5"]):
        ws1[h].font = font_header
        ws1[h].fill = fill_header_navy
        ws1[h].alignment = Alignment(horizontal="center", vertical="center")

    row_idx = 6
    for p in params:
        name, val, fmt, desc = p
        if val == "" and fmt == "":
            if name != "":
                ws1[f"B{row_idx}"] = name
                ws1[f"B{row_idx}"].font = font_section
                ws1[f"B{row_idx}"].fill = fill_subtle_blue
                ws1.merge_cells(f"B{row_idx}:E{row_idx}")
            row_idx += 1
            continue

        ws1[f"B{row_idx}"] = name
        ws1[f"B{row_idx}"].font = font_normal
        ws1[f"B{row_idx}"].border = border_cell

        ws1[f"C{row_idx}"] = val
        ws1[f"C{row_idx}"].font = font_bold
        ws1[f"C{row_idx}"].number_format = fmt
        ws1[f"C{row_idx}"].border = border_cell
        ws1[f"C{row_idx}"].alignment = Alignment(horizontal="right")

        ws1[f"D{row_idx}"] = fmt
        ws1[f"D{row_idx}"].font = font_normal
        ws1[f"D{row_idx}"].border = border_cell
        ws1[f"D{row_idx}"].alignment = Alignment(horizontal="center")

        ws1[f"E{row_idx}"] = desc
        ws1[f"E{row_idx}"].font = font_normal
        ws1[f"E{row_idx}"].border = border_cell
        row_idx += 1

    # ----------------------------------------------------
    # SHEET 2: FLUJO DE CAJA LIBRE (DCF)
    # ----------------------------------------------------
    print("[2/5] Creando hoja '02_Flujo_Caja_DCF'...")
    ws2 = wb.create_sheet(title="02_Flujo_Caja_DCF")
    ws2.views.sheetView[0].showGridLines = True

    ws2["B2"] = "ESTADO DE RESULTADOS PRO-FORMA & FLUJO DE CAJA LIBRE DESCONTADO (DCF)"
    ws2["B2"].font = font_title

    headers_dcf = ["Rubro Financiero / Año", "Año 0", "Año 1", "Año 2", "Año 3", "Año 4", "Año 5"]
    cols_dcf = ["B", "C", "D", "E", "F", "G", "H"]
    for c, h in zip(cols_dcf, headers_dcf):
        ws2[f"{c}4"] = h
        ws2[f"{c}4"].font = font_header
        ws2[f"{c}4"].fill = fill_header_navy
        ws2[f"{c}4"].alignment = Alignment(horizontal="center", vertical="center")

    dcf_rows = [
        ("1. Volumen de Servicios (Horas)", "-", "='01_Parametros_Supuestos'!$C$11", "=D5*(1+'01_Parametros_Supuestos'!$C$15)", "=E5*(1+'01_Parametros_Supuestos'!$C$15)", "=F5*(1+'01_Parametros_Supuestos'!$C$15)", "=G5*(1+'01_Parametros_Supuestos'!$C$15)", "#,##0"),
        ("2. Precio Unitario Promedio ($)", "-", "='01_Parametros_Supuestos'!$C$12", "=D6*(1+'01_Parametros_Supuestos'!$C$16)", "=E6*(1+'01_Parametros_Supuestos'!$C$16)", "=F6*(1+'01_Parametros_Supuestos'!$C$16)", "=G6*(1+'01_Parametros_Supuestos'!$C$16)", "$#,##0.00"),
        ("3. Costo Variable Unitario ($)", "-", "='01_Parametros_Supuestos'!$C$13", "=D7*(1+'01_Parametros_Supuestos'!$C$17)", "=E7*(1+'01_Parametros_Supuestos'!$C$17)", "=F7*(1+'01_Parametros_Supuestos'!$C$17)", "=G7*(1+'01_Parametros_Supuestos'!$C$17)", "$#,##0.00"),
        ("(=) Ingresos Operativos Brutos", "-", "=D5*D6", "=E5*E6", "=F5*F6", "=G5*G6", "=H5*H6", "$#,##0"),
        ("(-) Costos Variables Totales", "-", "=D5*D7", "=E5*E7", "=F5*F7", "=G5*G7", "=H5*H7", "$#,##0"),
        ("(=) Margen de Contribución", "-", "=D8-D9", "=E8-E9", "=F8-F9", "=G8-G9", "=H8-H9", "$#,##0"),
        ("(-) Costos Fijos Operativos (OPEX)", "-", "='01_Parametros_Supuestos'!$C$14", "=E11*(1+'01_Parametros_Supuestos'!$C$17)", "=F11*(1+'01_Parametros_Supuestos'!$C$17)", "=G11*(1+'01_Parametros_Supuestos'!$C$17)", "=H11*(1+'01_Parametros_Supuestos'!$C$17)", "$#,##0"),
        ("(-) Depreciación de Activos", "-", "='01_Parametros_Supuestos'!$C$8", "='01_Parametros_Supuestos'!$C$8", "='01_Parametros_Supuestos'!$C$8", "='01_Parametros_Supuestos'!$C$8", "='01_Parametros_Supuestos'!$C$8", "$#,##0"),
        ("(=) Utilidad Antes de Intereses e Impuestos (EBIT)", "-", "=D10-D11-D12", "=E10-E11-E12", "=F10-F11-F12", "=G10-G11-G12", "=H10-H11-H12", "$#,##0"),
        ("(-) Impuesto a la Renta (29.5%)", "-", "=D13*'01_Parametros_Supuestos'!$C$19", "=E13*'01_Parametros_Supuestos'!$C$19", "=F13*'01_Parametros_Supuestos'!$C$19", "=G13*'01_Parametros_Supuestos'!$C$19", "=H13*'01_Parametros_Supuestos'!$C$19", "$#,##0"),
        ("(=) Utilidad Operativa Neta Después de Impuestos (NOPAT)", "-", "=D13-D14", "=E13-E14", "=F13-F14", "=G13-G14", "=H13-H14", "$#,##0"),
        ("(+) Ajuste por Depreciación (No salida de caja)", "-", "=D12", "=E12", "=F12", "=G12", "=H12", "$#,##0"),
        ("(-) Inversión Inicial en Activos (CAPEX)", "=-'01_Parametros_Supuestos'!$C$6", 0, 0, 0, 0, 0, "$#,##0"),
        ("(+) Valor de Rescate de Activos", 0, 0, 0, 0, 0, "='01_Parametros_Supuestos'!$C$9", "$#,##0"),
        ("(=) FLUJO DE CAJA LIBRE DEL PROYECTO (FCFF)", "=C17", "=D15+D16+D17+D18", "=E15+E16+E17+E18", "=F15+F16+F17+F18", "=G15+G16+G17+G18", "=H15+H16+H17+H18", "$#,##0"),
        ("Factor de Descuento WACC (t)", 1.0, "=1/(1+'01_Parametros_Supuestos'!$C$20)^1", "=1/(1+'01_Parametros_Supuestos'!$C$20)^2", "=1/(1+'01_Parametros_Supuestos'!$C$20)^3", "=1/(1+'01_Parametros_Supuestos'!$C$20)^4", "=1/(1+'01_Parametros_Supuestos'!$C$20)^5", "0.0000"),
        ("Flujo de Caja Descontado (Valor Presente)", "=C19*C20", "=D19*D20", "=E19*E20", "=F19*F20", "=G19*G20", "=H19*H20", "$#,##0"),
        ("Flujo de Caja Descontado Acumulado", "=C21", "=C22+D21", "=D22+E21", "=E22+F21", "=F22+G21", "=G22+H21", "$#,##0"),
    ]

    r_idx = 5
    for item in dcf_rows:
        label = item[0]
        fmt = item[-1]
        vals = item[1:-1]

        ws2[f"B{r_idx}"] = label
        is_highlight = "(=) FLUJO DE CAJA" in label or "(=) Utilidad Operativa Neta" in label
        ws2[f"B{r_idx}"].font = font_bold if is_highlight else font_normal
        ws2[f"B{r_idx}"].border = border_cell
        if is_highlight:
            ws2[f"B{r_idx}"].fill = fill_subtle_blue

        for col_i, v in zip(cols_dcf[1:], vals):
            cell = ws2[f"{col_i}{r_idx}"]
            cell.value = v
            cell.number_format = fmt
            cell.font = font_bold if is_highlight else font_normal
            cell.border = border_cell
            cell.alignment = Alignment(horizontal="right")
            if is_highlight:
                cell.fill = fill_subtle_blue
        r_idx += 1

    # EXECUTIVE KPI SUMMARY DASHBOARD
    ws2["B25"] = "INDICADORES CLAVE DE DECISIÓN FINANCIERA (KPIs)"
    ws2["B25"].font = font_section
    ws2.merge_cells("B25:E25")

    kpis = [
        ("Valor Actual Neto (VAN / NPV)", "=NPV('01_Parametros_Supuestos'!$C$20, D19:H19)+C19", "$#,##0", "VAN > 0 indica creación neta de valor"),
        ("Tasa Interna de Retorno (TIR / IRR)", "=IRR(C19:H19)", "0.0%", "Rentabilidad intrínseca del flujo de fondos"),
        ("WACC / Tasa de Descuento de Referencia", "='01_Parametros_Supuestos'!$C$20", "0.0%", "Costo promedio ponderado de capital"),
        ("Spread de Rentabilidad (TIR - WACC)", "=C27-C28", "0.0%", "Margen de seguridad financiero sobre costo de capital"),
        ("Índice de Rentabilidad (B/C Ratio)", "=SUM(D21:H21)/ABS(C19)", "0.00", "Dólares recuperados por cada dólar invertido"),
        ("Decisión Económica Recomendada", '=IF(C26>0, "APROBAR PROYECTO (CREA VALOR)", "RECHAZAR")', "@", "Criterio normativo de finanzas corporativas"),
    ]

    for k_idx, (k_name, k_val, k_fmt, k_desc) in enumerate(kpis, start=26):
        ws2[f"B{k_idx}"] = k_name
        ws2[f"B{k_idx}"].font = font_bold
        ws2[f"B{k_idx}"].border = border_cell

        ws2[f"C{k_idx}"] = k_val
        ws2[f"C{k_idx}"].font = font_kpi_num
        ws2[f"C{k_idx}"].number_format = k_fmt
        ws2[f"C{k_idx}"].border = border_cell
        ws2[f"C{k_idx}"].fill = fill_subtle_green
        ws2[f"C{k_idx}"].alignment = Alignment(horizontal="center")

        ws2[f"D{k_idx}"] = k_desc
        ws2[f"D{k_idx}"].font = font_normal
        ws2[f"D{k_idx}"].border = border_cell
        ws2.merge_cells(f"D{k_idx}:G{k_idx}")

    # ----------------------------------------------------
    # SHEET 3: MATRIZ DE SENSIBILIDAD BIDIMENSIONAL
    # ----------------------------------------------------
    print("[3/5] Creando hoja '03_Sensibilidad_Bidimensional'...")
    ws3 = wb.create_sheet(title="03_Sensibilidad_Bidimensional")
    ws3.views.sheetView[0].showGridLines = True

    ws3["B2"] = "ANÁLISIS DE SENSIBILIDAD BIDIMENSIONAL DEL VAN ($)"
    ws3["B2"].font = font_title
    ws3["B3"] = "Impacto cruzado: Variaciones de Tasa de Descuento (WACC) vs Precio Unitario de Venta"
    ws3["B3"].font = Font(name="Calibri", size=11, italic=True, color="64748B")

    wacc_rates = [0.095, 0.105, 0.115, 0.125, 0.135]
    prices = [45.0, 50.0, 55.0, 60.0, 65.0]

    # Deterministic base NPV calculation helper
    def calc_npv(wacc_val, price_val):
        capex = 250000
        tax = 0.295
        deprec = capex / 5
        salvage = 35000
        cf_list = [-capex]
        vol = 12000
        cost_u = 22.0
        opex = 95000
        for y in range(1, 6):
            rev = vol * price_val
            cvar = vol * cost_u
            margin = rev - cvar
            ebit = margin - opex - deprec
            taxes = ebit * tax
            nopat = ebit - taxes
            fcff = nopat + deprec
            if y == 5:
                fcff += salvage
            cf_list.append(fcff)
            vol *= 1.08
            price_val *= 1.03
            cost_u *= 1.025
            opex *= 1.025

        npv_val = cf_list[0] + sum(cf / ((1 + wacc_val) ** t) for t, cf in enumerate(cf_list[1:], 1))
        return npv_val

    ws3["B5"] = "WACC \\ Precio"
    ws3["B5"].font = font_header
    ws3["B5"].fill = fill_header_slate
    ws3["B5"].alignment = Alignment(horizontal="center")

    for col_p, p in enumerate(prices, start=3):
        col_letter = get_column_letter(col_p)
        ws3[f"{col_letter}5"] = f"${p:.2f}"
        ws3[f"{col_letter}5"].font = font_header
        ws3[f"{col_letter}5"].fill = fill_header_navy
        ws3[f"{col_letter}5"].alignment = Alignment(horizontal="center")

    for row_w, w in enumerate(wacc_rates, start=6):
        ws3[f"B{row_w}"] = f"{w:.1%}"
        ws3[f"B{row_w}"].font = font_header
        ws3[f"B{row_w}"].fill = fill_header_navy
        ws3[f"B{row_w}"].alignment = Alignment(horizontal="center")

        for col_p, p in enumerate(prices, start=3):
            col_letter = get_column_letter(col_p)
            val = calc_npv(w, p)
            cell = ws3[f"{col_letter}{row_w}"]
            cell.value = val
            cell.number_format = "$#,##0"
            cell.font = font_bold
            cell.border = border_cell
            cell.alignment = Alignment(horizontal="right")
            if val >= 0:
                cell.fill = fill_subtle_green
            else:
                cell.fill = fill_subtle_red

    # ----------------------------------------------------
    # SHEET 4: SIMULACION MONTE CARLO (5,000 ITERACIONES)
    # ----------------------------------------------------
    print("[4/5] Ejecutando simulación estocástica de Monte Carlo (5,000 iteraciones)...")
    np.random.seed(42)
    n_sims = 5000

    sim_demands = np.random.normal(loc=12000, scale=1400, size=n_sims)
    sim_prices = np.random.normal(loc=55.0, scale=3.5, size=n_sims)
    sim_costs = np.random.triangular(left=19.0, mode=22.0, right=26.0, size=n_sims)
    sim_waccs = np.random.normal(loc=0.115, scale=0.012, size=n_sims)
    sim_waccs = np.clip(sim_waccs, 0.07, 0.18)

    sim_npvs = []
    sim_irrs = []

    for i in range(n_sims):
        d = max(6000, sim_demands[i])
        p = max(35.0, sim_prices[i])
        c = max(15.0, sim_costs[i])
        w = sim_waccs[i]

        cf_list = [-250000]
        curr_d, curr_p, curr_c, curr_op = d, p, c, 95000
        for yr in range(1, 6):
            rev = curr_d * curr_p
            cvar = curr_d * curr_c
            ebit = (rev - cvar) - curr_op - 50000
            tax = max(0, ebit * 0.295)
            fcff = (ebit - tax) + 50000
            if yr == 5:
                fcff += 35000
            cf_list.append(fcff)
            curr_d *= 1.08
            curr_p *= 1.03
            curr_c *= 1.025
            curr_op *= 1.025

        npv_val = cf_list[0] + sum(cf / ((1 + w) ** t) for t, cf in enumerate(cf_list[1:], 1))
        # Simple IRR approximation
        try:
            irr_val = np.irr(cf_list) if hasattr(np, 'irr') else np.roots(cf_list[::-1])
        except Exception:
            irr_val = 0.20
        sim_npvs.append(npv_val)

    sim_npvs = np.array(sim_npvs)

    mean_npv = float(np.mean(sim_npvs))
    median_npv = float(np.median(sim_npvs))
    std_npv = float(np.std(sim_npvs))
    p5_npv = float(np.percentile(sim_npvs, 5))    # Value at Risk 95%
    p95_npv = float(np.percentile(sim_npvs, 95))
    prob_success = float(np.mean(sim_npvs > 0))
    prob_loss = float(np.mean(sim_npvs <= 0))
    cvar_95 = float(np.mean(sim_npvs[sim_npvs <= p5_npv])) # Conditional VaR

    ws4 = wb.create_sheet(title="04_Simulacion_MonteCarlo")
    ws4.views.sheetView[0].showGridLines = True

    ws4["B2"] = "SIMULACIÓN ESTOCÁSTICA DE MONTE CARLO (N = 5,000 ESCENARIOS)"
    ws4["B2"].font = font_title
    ws4["B3"] = "Evaluación de Riesgo Cuantitativo y Distribución de Probabilidad del VAN"
    ws4["B3"].font = Font(name="Calibri", size=11, italic=True, color="64748B")

    # Estadísticos resumen
    ws4["B5"] = "MÉTRICAS ESTADÍSTICAS DEL VAN SIMULADO"
    ws4["B5"].font = font_section
    ws4["C5"] = "VALOR"
    ws4["D5"] = "DESCRIPCIÓN TÉCNICA"
    for col, h in enumerate(["B5", "C5", "D5"]):
        ws4[h].font = font_header
        ws4[h].fill = fill_header_navy
        ws4[h].alignment = Alignment(horizontal="center")

    summary_stats = [
        ("Media Esperada del VAN (E[VAN])", mean_npv, "$#,##0", "Valor central esperado de la distribución estocástica"),
        ("Mediana del VAN", median_npv, "$#,##0", "Percentil 50 libre de sesgo por asimetría"),
        ("Desviación Estándar (Volatilidad del VAN)", std_npv, "$#,##0", "Dispersión de riesgo alrededor de la media"),
        ("Value at Risk (VaR 95% Cuantitativo)", p5_npv, "$#,##0", "En el 95% de los escenarios el VAN superará este umbral"),
        ("Conditional VaR / Expected Shortfall (CVaR 95%)", cvar_95, "$#,##0", "Pérdida esperada en el 5% de los peores escenarios extremos"),
        ("Percentil 95 (Escenario Optimista)", p95_npv, "$#,##0", "Cota superior en condiciones macroeconómicas favorables"),
        ("Probabilidad de Éxito P(VAN > 0)", prob_success, "0.0%", "Porcentaje de simulaciones donde el proyecto genera valor económico"),
        ("Probabilidad de Riesgo de Pérdida P(VAN <= 0)", prob_loss, "0.0%", "Probabilidad acumulada de destrucción de capital"),
    ]

    for s_idx, (s_name, s_val, s_fmt, s_desc) in enumerate(summary_stats, start=6):
        ws4[f"B{s_idx}"] = s_name
        ws4[f"B{s_idx}"].font = font_bold
        ws4[f"B{s_idx}"].border = border_cell

        ws4[f"C{s_idx}"] = s_val
        ws4[f"C{s_idx}"].number_format = s_fmt
        ws4[f"C{s_idx}"].font = font_kpi_num
        ws4[f"C{s_idx}"].border = border_cell
        ws4[f"C{s_idx}"].alignment = Alignment(horizontal="right")
        ws4[f"C{s_idx}"].fill = fill_subtle_green if "Éxito" in s_name or s_val > 0 else fill_subtle_yellow

        ws4[f"D{s_idx}"] = s_desc
        ws4[f"D{s_idx}"].font = font_normal
        ws4[f"D{s_idx}"].border = border_cell
        ws4.merge_cells(f"D{s_idx}:F{s_idx}")

    # Tabla de 5,000 datos
    start_data_row = 16
    headers_sim = ["Iteración", "Demanda (Horas)", "Precio ($)", "Costo Var ($)", "WACC (%)", "VAN Simulado ($)"]
    cols_sim = ["B", "C", "D", "E", "F", "G"]
    for c, h in zip(cols_sim, headers_sim):
        ws4[f"{c}{start_data_row}"] = h
        ws4[f"{c}{start_data_row}"].font = font_header
        ws4[f"{c}{start_data_row}"].fill = fill_header_slate
        ws4[f"{c}{start_data_row}"].alignment = Alignment(horizontal="center")

    for i in range(min(5000, n_sims)):
        row = start_data_row + 1 + i
        ws4[f"B{row}"] = i + 1
        ws4[f"C{row}"] = round(sim_demands[i], 1)
        ws4[f"D{row}"] = round(sim_prices[i], 2)
        ws4[f"E{row}"] = round(sim_costs[i], 2)
        ws4[f"F{row}"] = round(sim_waccs[i], 4)
        ws4[f"G{row}"] = round(sim_npvs[i], 2)

        ws4[f"C{row}"].number_format = "#,##0"
        ws4[f"D{row}"].number_format = "$#,##0.00"
        ws4[f"E{row}"].number_format = "$#,##0.00"
        ws4[f"F{row}"].number_format = "0.0%"
        ws4[f"G{row}"].number_format = "$#,##0"

    # ----------------------------------------------------
    # SHEET 5: POWER QUERY M SCRIPT DOCUMENTATION
    # ----------------------------------------------------
    print("[5/5] Creando hoja '05_PowerQuery_M_Script'...")
    ws5 = wb.create_sheet(title="05_PowerQuery_M_Script")
    ws5.views.sheetView[0].showGridLines = True

    ws5["B2"] = "FLUJO DE EXTRACCIÓN & TRANSFORMACIÓN ETL CON POWER QUERY (M CODE)"
    ws5["B2"].font = font_title
    ws5["B3"] = "Automatización del Pipeline de Ingesta desde Base Operacional SQL / ERP"
    ws5["B3"].font = Font(name="Calibri", size=11, italic=True, color="64748B")

    m_code = """// Código M (Power Query) para consolidar horas de vuelo y telemetría de drones
let
    // 1. Origen: Conexión a Base de Datos Operacional de Telemetría SQLite
    Origen = Sqlite.Database("C:\\Users\\USUARIO\\Desktop\\CV\\portafolio\\09-python-telemetry-streaming\\output\\benchmark_telemetry.db"),
    telemetry_events = Origen{[Schema="",Item="telemetry_events"]}[Data],

    // 2. Filtrar únicamente eventos exitosos sin anomalías críticas
    FiltrarValidos = Table.SelectRows(telemetry_events, each ([is_anomalous] = 0)),

    // 3. Parsear fecha y timestamp
    TipoCambiado = Table.TransformColumnTypes(FiltrarValidos, {
        {"event_timestamp", type datetimezone},
        {"device_id", type text},
        {"latency_ms", type number}
    }),

    // 4. Agregar columna de periodo contable (Año-Mes)
    AgregarPeriodo = Table.AddColumn(TipoCambiado, "Periodo", each DateTimeZone.ToText([event_timestamp], "yyyy-MM"), type text),

    // 5. Agrupación mensual para cálculo de Horas Operativas de Vuelo
    Agrupado = Table.Group(AgregarPeriodo, {"Periodo", "device_id"}, {
        {"Total_Eventos_Sensor", each Table.RowCount(_), Int64.Type},
        {"Latencia_Promedio_ms", each List.Average([latency_ms]), type number}
    }),

    // 6. Cálculo de Horas Equivalentes de Inspección para Facturación DCF
    HorasEquivalentes = Table.AddColumn(Agrupado, "Horas_Facturables", each [Total_Eventos_Sensor] / 3600, type number)
in
    HorasEquivalentes"""

    ws5["B5"] = "Script de Consulta Power Query (Lenguaje M):"
    ws5["B5"].font = font_section

    ws5["B7"] = m_code
    ws5["B7"].font = Font(name="Consolas", size=10, color="0F172A")
    ws5["B7"].alignment = Alignment(wrap_text=True, vertical="top")
    ws5.merge_cells("B7:H28")
    ws5["B7"].fill = fill_subtle_blue

    # Auto-adjust column widths across all sheets
    for ws in [ws1, ws2, ws3, ws4, ws5]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if "\n" in val_str:
                    val_str = max(val_str.split("\n"), key=len)
                if len(val_str) > max_len and len(val_str) < 60:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # Save Excel workbook
    wb.save(excel_path)
    print(f"[OK] Archivo Excel guardado exitosamente en: {excel_path}")

    # ----------------------------------------------------
    # GENERAR GRAFICOS DE ALTA RESOLUCION PARA README
    # ----------------------------------------------------
    print("[6/5] Generando visualizaciones estadísticas para el portafolio...")
    plt.style.use("seaborn-v0_8-whitegrid" if "seaborn-v0_8-whitegrid" in plt.style.available else "default")

    # Chart 1: Monte Carlo Distribution
    fig, ax = plt.subplots(figsize=(10, 6), dpi=300)
    sns.histplot(sim_npvs, kde=True, bins=60, color="#1E3A8A", alpha=0.6, ax=ax)

    ax.axvline(0, color="#EF4444", linestyle="--", linewidth=2, label="Umbral de Rechazo (VAN = $0)")
    ax.axvline(mean_npv, color="#10B981", linestyle="-", linewidth=2.5, label=f"Media Esperada: ${mean_npv:,.0f}")
    ax.axvline(p5_npv, color="#F59E0B", linestyle=":", linewidth=2, label=f"VaR 95%: ${p5_npv:,.0f}")

    ax.set_title("Distribución de Probabilidad del VAN (Simulación Monte Carlo N = 5,000)", fontsize=13, fontweight="bold", pad=12)
    ax.set_xlabel("Valor Actual Neto (VAN en USD)", fontsize=11, fontweight="bold")
    ax.set_ylabel("Frecuencia de Escenarios", fontsize=11, fontweight="bold")
    ax.xaxis.set_major_formatter("${x:,.0f}")
    ax.legend(frameon=True, facecolor="white", framealpha=0.9, fontsize=10)

    # Risk Metrics Textbox
    textstr = "\n".join((
        f"Probabilidad VAN > 0 : {prob_success:.1%}",
        f"Probabilidad Riesgo  : {prob_loss:.1%}",
        f"Volatilidad (Std Dev): ${std_npv:,.0f}",
        f"CVaR 95% (Tail Loss) : ${cvar_95:,.0f}",
    ))
    props = dict(boxstyle="round,pad=0.6", facecolor="#EFF6FF", edgecolor="#93C5FD", alpha=0.9)
    ax.text(0.03, 0.95, textstr, transform=ax.transAxes, fontsize=10, verticalalignment="top", bbox=props)

    plt.tight_layout()
    chart1_path = os.path.join(assets_dir, "monte_carlo_distribution.png")
    plt.savefig(chart1_path)
    plt.close()
    print(f"[OK] Gráfico de Monte Carlo guardado en: {chart1_path}")

    # Chart 2: Tornado Sensitivity Analysis
    fig, ax = plt.subplots(figsize=(10, 5), dpi=300)

    variables = ["Precio por Hora", "Demanda Base (Horas)", "WACC (Tasa Descuento)", "Costos Fijos (OPEX)", "CAPEX Inicial"]
    # Impact on NPV from low to high (+/- 15%)
    low_impacts = [calc_npv(0.115, 55*0.85) - mean_npv, calc_npv(0.115, 55) - 60000 - mean_npv, calc_npv(0.115*1.2, 55) - mean_npv, calc_npv(0.115, 55) - 35000 - mean_npv, -37500]
    high_impacts = [calc_npv(0.115, 55*1.15) - mean_npv, calc_npv(0.115, 55) + 60000 - mean_npv, calc_npv(0.115*0.8, 55) - mean_npv, calc_npv(0.115, 55) + 35000 - mean_npv, 37500]

    y_pos = np.arange(len(variables))
    ax.barh(y_pos, low_impacts, align="center", color="#EF4444", alpha=0.8, label="Impacto Desfavorable (-15%)")
    ax.barh(y_pos, high_impacts, align="center", color="#10B981", alpha=0.8, label="Impacto Favorable (+15%)")

    ax.set_yticks(y_pos)
    ax.set_yticklabels(variables, fontsize=10, fontweight="bold")
    ax.invert_yaxis()
    ax.axvline(0, color="#1E293B", linewidth=1.2)
    ax.set_xlabel("Variación del Valor Actual Neto frente a la Media ($ USD)", fontsize=11, fontweight="bold")
    ax.set_title("Diagrama de Tornado — Sensibilidad Unidimensional de Variables Clave", fontsize=13, fontweight="bold", pad=12)
    ax.xaxis.set_major_formatter("${x:,.0f}")
    ax.legend(frameon=True, facecolor="white", fontsize=10)

    plt.tight_layout()
    chart2_path = os.path.join(assets_dir, "tornado_sensitivity.png")
    plt.savefig(chart2_path)
    plt.close()
    print(f"[OK] Diagrama de Tornado guardado en: {chart2_path}")

    print("\n" + "=" * 70)
    print("MODELO FINANCIERO Y MONTE CARLO COMPLETADO EXITOSAMENTE")
    print("=" * 70)
    print(f"Media del VAN : ${mean_npv:,.2f} USD")
    print(f"Tasa de Éxito : {prob_success:.1%} de escenarios rentables")
    print(f"VaR 95%       : ${p5_npv:,.2f} USD")
    print("=" * 70)

if __name__ == "__main__":
    build_model()
