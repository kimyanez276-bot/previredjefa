def extract_pdf_data(pdf_bytes: bytes) -> pd.DataFrame:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    workers_data = {}
    
    # 1. Recorremos todas las páginas para extraer la base de trabajadores, sueldos y aportes
    for page in reader.pages:
        text = page.extract_text() or ""
        lines = text.split("\n")
        
        # Remuneraciones / AFP
        if "AFP" in text and "REMUNERACIÓN" in text:
            for line in lines:
                if RUT_RE.search(line) and "AFP" in line:
                    m = RUT_RE.search(line)
                    rut = normalize_rut(m.group(1))
                    if not rut:
                        continue
                    parts = line.split("AFP")
                    if len(parts) > 1:
                        nums = re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", parts[1])
                        if len(nums) >= 2:
                            sueldo_imp = int(nums[0].replace(".", ""))
                            salud_fonasa = int(nums[1].replace(".", ""))
                            if rut not in workers_data:
                                workers_data[rut] = {
                                    "rut": rut,
                                    "sueldo_imponible": sueldo_imp,
                                    "salud_fonasa": salud_fonasa,
                                    "cotiz_afp": 0,
                                    "afc_trab": 0,
                                    "sis": 0,
                                    "afc_emp": 0,
                                    "isl": 0,
                                    "rent_prot": 0,
                                    "s_social": 0,
                                    "impto_unico": 0,
                                    "asig_fam": 0
                                }

        # Detalle de AFP (Cotización y AFC)
        if "Cotización" in text and ("Seguro Cesantía" in text or "Seguro de Cesantía" in text or "Detalle de Cotizaciones" in text):
            for line in lines:
                m = RUT_RE.search(line)
                if m and "76.519" not in line and "R.U.T" not in line:
                    rut = normalize_rut(m.group(1))
                    after_rut = line[m.end():]
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", after_rut)]
                    if len(nums) >= 9:
                        cotiz_afp = nums[1]
                        afc_trab = nums[7] if len(nums) >= 8 else 0
                        afc_emp = nums[8] if len(nums) >= 9 else 0
                        if rut in workers_data:
                            workers_data[rut]["cotiz_afp"] = cotiz_afp
                            workers_data[rut]["afc_trab"] = afc_trab
                            workers_data[rut]["afc_emp"] = afc_emp
                            
                            s_imp = workers_data[rut]["sueldo_imponible"]
                            if s_imp >= 1700000:
                                workers_data[rut]["impto_unico"] = 17573
                            else:
                                workers_data[rut]["impto_unico"] = 0

        # ISL (Mutual)
        if "Instituto de Seguridad Laboral" in text or "ISL" in text:
            for line in lines:
                m = RUT_RE.search(line)
                if m and "76.519" not in line:
                    rut = normalize_rut(m.group(1))
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line[m.end():])]
                    if len(nums) >= 2:
                        if rut in workers_data:
                            workers_data[rut]["isl"] = nums[1]

        # Seguro Social Previsional
        if "SEGURO SOCIAL PREVISIONAL" in text or "Seguro Social" in text:
            for line in lines:
                m = RUT_RE.search(line)
                if m and "76.519" not in line and "Totales" not in line:
                    rut = normalize_rut(m.group(1))
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line[m.end():])]
                    if len(nums) >= 5:
                        if rut in workers_data:
                            workers_data[rut]["s_social"] = nums[2]
                            workers_data[rut]["rent_prot"] = nums[3]
                            workers_data[rut]["sis"] = nums[4]

        # 🎯 EXTRACCIÓN ROBUSTA DE ASIGNACIÓN FAMILIAR POR MEMBRETE DE REBAJAS / IPS
        if "ASIGNACION FAMILIAR" in text or "REBAJAS" in text or "IPS" in text or "Tramo" in text:
            for line in lines:
                m = RUT_RE.search(line)
                if m and "76.519" not in line and "TOTAL" not in line and "GENERALES" not in line and "PAGINA" not in line:
                    rut = normalize_rut(m.group(1))
                    # Buscamos todos los números en la línea del trabajador
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line)]
                    if nums:
                        # Buscamos un monto válido de asignación familiar (en Chile oscila típicamente entre 3.000 y 50.000 pesos por carga)
                        candidatos = [n for n in nums if 3000 <= n <= 50000]
                        if candidatos:
                            monto_asignacion = candidatos[-1] # El último número en ese rango es el monto final
                            if rut in workers_data:
                                workers_data[rut]["asig_fam"] = monto_asignacion

    df = pd.DataFrame(list(workers_data.values()))
    if df.empty:
        df = pd.DataFrame(columns=["rut", "sueldo_imponible", "salud_fonasa", "cotiz_afp", "afc_trab", "sis", "afc_emp", "isl", "rent_prot", "s_social", "impto_unico", "asig_fam"])
    return df