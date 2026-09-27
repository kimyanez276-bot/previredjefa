def extract_asignacion_familiar(pdf_reader, workers_data):
    """
    Extrae la Asignación Familiar estrictamente desde la sección de rebajas/IPS de Previred.
    """
    section_keywords = ["ASIGNACION FAMILIAR", "REBAJAS", "IPS"]

    for page in pdf_reader.pages:
        text = page.extract_text()
        if not text:
            continue

        text_upper = text.upper()

        if not any(keyword in text_upper for keyword in section_keywords):
            continue

        lines = text.split("\n")

        for line in lines:
            rut_match = RUT_RE.search(line)
            if not rut_match:
                continue

            rut = normalize_rut(rut_match.group())

            # Buscamos números en la línea pero filtramos para que NO agarre RUTs ni montos de salud/imponibles grandes
            nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line)]
            
            # Los montos de asignación familiar en Chile son acotados (ej. entre 3.000 y 30.000 pesos por carga)
            # Excluimos explícitamente los números que parezcan sueldos, salud o partes de RUT.
            candidatos_asig = [n for n in nums if 3000 <= n <= 35000]

            if candidatos_asig:
                # El monto real de asignación familiar es el valor que corresponde a la columna Monto
                monto = candidatos_asig[-1]
                if rut in workers_data:
                    workers_data[rut]["asig_fam"] = monto
                else:
                    workers_data[rut] = {
                        "rut": rut,
                        "sueldo_imponible": 0,
                        "salud_fonasa": 0,
                        "cotiz_afp": 0,
                        "afc_trab": 0,
                        "sis": 0,
                        "afc_emp": 0,
                        "isl": 0,
                        "rent_prot": 0,
                        "s_social": 0,
                        "impto_unico": 0,
                        "asig_fam": monto
                    }

    return workers_data