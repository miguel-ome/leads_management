from config import FIELDS_TABLE_LEADS


class Filter_Service:
  @staticmethod
  def filter_data(self, filter_ui):
    resultado = []

    for row in filter_ui.dados:
        manter = True

        for i in range(len(FIELDS_TABLE_LEADS)):
            campo = FIELDS_TABLE_LEADS[i]
            dado = str(row[i]).lower()

            if campo == "Salário":
                operador, entry = filter_ui.entries_filtro[i]
                op = operador.get().strip()
                val = entry.get().strip()

                if val:
                    try:
                        # Verifica se o valor está vazio ou é None
                        if row[i] is None or row[i] == "":
                            manter = False
                            continue
                        
                        dado_float = float(row[i])
                        val_float = float(val)

                        if op == "=" and not (dado_float == val_float):
                            manter = False
                        elif op == ">" and not (dado_float > val_float):
                            manter = False
                        elif op == "<" and not (dado_float < val_float):
                            manter = False
                        elif op == ">=" and not (dado_float >= val_float):
                            manter = False
                        elif op == "<=" and not (dado_float <= val_float):
                            manter = False
                    except ValueError:
                        manter = False
            else:
                filtro = filter_ui.entries_filtro[i].get().lower()
                if filtro and filtro not in dado:
                    manter = False

        if manter:
            resultado.append(row)

    filter_ui.atualizar_tabela(resultado)
