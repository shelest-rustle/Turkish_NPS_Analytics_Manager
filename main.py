import pandas
import string

from config import *

class Main():

    def __init__(self, origin_analytic):
        self.origin_analytic = origin_analytic
        self.origin_df = pandas.read_excel(origin_analytic, dtype=str)
        self.result_df = pandas.DataFrame()
        self.tum_data = pandas.DataFrame()
        self.tum_data_tekil = pandas.DataFrame()
        self.ulasilan_data = pandas.DataFrame()
        self.basarili = pandas.DataFrame()
        self.ozet = pandas.DataFrame(index=range(10), columns=range(10))
        self.pivot = pandas.DataFrame(index=range(200), columns=range(200))
        self.grafikler = pandas.DataFrame(index=range(500), columns=range(500))
        self.unit_unique_values_urun_grubu = {}
        self.unit_unique_values_satis_kanali = {}
        self.units = UNITS

    def make_tum_data(self):
        self.origin_df['msisdn'] = self.origin_df['msisdn']
        self.tum_data["msisdn"] = self.origin_df["msisdn"]
        self.tum_data["MARKA"] = self.origin_df["MARKA"]
        self.tum_data["call_status"] = self.origin_df["call_status"]
        self.tum_data["result"] = self.origin_df["result"]
        # result на турецком
        self.tum_data["Anket sonucu"] = self.origin_df["Anket sonucu"]
        # дозвон/недозвон - ULAŞILDI/ULAŞILAMADI
        self.tum_data["Ulaşma Sonuç Kodu"] = self.origin_df["Ulaşma Sonuç Kodu"]
        # проставляется по сип ответу + по некоторым звонкам 200ок(недозвон)
        self.tum_data["ALT KOD"] = self.origin_df["ALT KOD"]
        self.tum_data["Satın alma kanalı"] = self.origin_df["Satın alma kanalı"]
        self.tum_data["TEKNOCLUB_SEGMENT"] = self.origin_df["TEKNOCLUB_SEGMENT"]
        self.tum_data["Ürün grubu"] = self.origin_df["Ürün grubu"]
        self.tum_data["duration"] = self.origin_df["duration"]
        self.tum_data["get_duration"] = self.origin_df["get_duration"]
        self.tum_data["call_start_time"] = self.origin_df["call_start_time"]
        self.tum_data["attempt"] = self.origin_df["attempt"]
        self.tum_data["call_transcript"] = self.origin_df["call_transcript"]

        self.tum_data["delivery"] = self.origin_df["delivery"]
        # delivery на  турецком
        self.tum_data["Kargo teslimat"] = self.origin_df["Kargo teslimat"]

        self.tum_data["installation"] = self.origin_df["installation"]
        # installation на  турецком
        self.tum_data["Ürün kurulumu"] = self.origin_df["Ürün kurulumu"]

        self.tum_data["recall"] = self.origin_df["recall"]

        self.tum_data["deadline"] = self.origin_df["deadline"]
        # deadline на  турецком
        self.tum_data["Kurulum için zamanında Sonuç"] = self.origin_df["Kurulum için zamanında Sonuç"]

        self.tum_data["introduction"] = self.origin_df["introduction"]
        # introduction на  турецком
        self.tum_data["Ürün tanıtımı"] = self.origin_df["Ürün tanıtımı"]

        self.tum_data["attitude"] = self.origin_df["attitude"]
        # attitude на  турецком
        self.tum_data["Servis tutum Sonuç"] = self.origin_df["Servis tutum Sonuç"]

        self.tum_data["product"] = self.origin_df["product"]
        # product на  турецком
        self.tum_data["Üründen mennunihyet Sonuç"] = self.origin_df["Üründen mennunihyet Sonuç"]

    def make_tum_data_tekil(self):
        self.tum_data["dialog_uuid"] = self.origin_df["dialog_uuid"]
        self.tum_data_tekil = (self.tum_data[self.tum_data.duplicated(subset="msisdn", keep=False)]).groupby('msisdn').filter(lambda x: x['dialog_uuid'].nunique() > 1)
        self.tum_data.drop("dialog_uuid", axis=1, inplace=True)

    def make_ulasilan_data(self):
        self.ulasilan_data = self.tum_data[self.tum_data["call_status"] == "200 OK"]

    def make_basarili(self):
        # Anket sonucu == BAŞARILI <...>
        self.basarili = self.tum_data[self.tum_data["Anket sonucu"].str.contains("BAŞARILI", na=False)]

    def excel_to_index(self, cell):
        col_letter = ''.join(filter(str.isalpha, cell))
        row_number = ''.join(filter(str.isdigit, cell))
        col_index = string.ascii_uppercase.index(col_letter.upper())
        row_index = int(row_number) - 1
        return row_index, col_index

    def make_ozet(self):
        name_fields = {
            "B2": "TOPLAM ARAMA",
            "B3": "TOPLAM ARAMA TEKİL",
            "B4": "ULAŞILAN DATA",
            "B5": "ULAŞILAN DATA TEKİL",
            "B6": "ULAŞMA ORANI",
            "B7": "ULAŞMA ORANI TEKİL",
            "B8": "BAŞARILI ANKET",
            "B9": "BAŞARILI ANKET %",
            "E2": "ULAŞILAMADI",
            "E3": "GÖRÜŞME İÇERİĞİ MEVCUT DEĞİL",
            "E4": "KULLANILMAYAN NUMARA",
            "E5": "MEŞGUL",
            "E6": "MÜKERRER ARAMA",
            "E7": "NUMARA EKSİK-HATALI",
            "E8": "TEKNİK SORUN",
            "E9": "TELESEKRETER",
        }

        for cell, value in name_fields.items():
            row, col = self.excel_to_index(cell)
            self.ozet.iloc[row, col] = value

        row, col = self.excel_to_index("C2")
        toplam_arama = self.tum_data.shape[0]
        toplam_arama_tekil = self.tum_data_tekil.shape[0]
        ulasilan_data = self.tum_data[self.tum_data["Ulaşma Sonuç Kodu"] == "ULAŞILDI"].shape[0]
        ulasilan_data_tekil = self.tum_data_tekil[self.tum_data_tekil["Ulaşma Sonuç Kodu"] == "ULAŞILDI"].shape[0]
        self.ozet.iloc[row, col] = toplam_arama  # C2 TOPLAM ARAMA
        self.ozet.iloc[row+1, col] = toplam_arama_tekil  # C3 TOPLAM ARAMA TEKİL
        self.ozet.iloc[row+2, col] = ulasilan_data  # C4 ULAŞILAN DATA
        self.ozet.iloc[row+3, col] = ulasilan_data_tekil  # C5 ULAŞILAN DATA TEKİL
        try:
            percent = ulasilan_data / toplam_arama * 100
        except ZeroDivisionError:
            percent = 0
        self.ozet.iloc[row+4, col] = f"{percent:.2f}%"  # C6 ULAŞMA ORANI
        try:
            percent = ulasilan_data_tekil / toplam_arama_tekil * 100
        except ZeroDivisionError:
            percent = 0
        self.ozet.iloc[row+5, col] = f"{percent:.2f}%"  # C7 ULAŞMA ORANI TEKİL
        basarili_anket = self.basarili.shape[0]
        self.ozet.iloc[row+6, col] = basarili_anket  # C8 BAŞARILI ANKET
        try:
            percent = basarili_anket / ulasilan_data * 100
        except ZeroDivisionError:
            percent = 0
        self.ozet.iloc[row+7, col] = f"{percent:.2f}%"  # C9 BAŞARILI ANKET %

        row, col = self.excel_to_index("F2")
        ulasilamadi = self.tum_data[self.tum_data["Ulaşma Sonuç Kodu"] == "ULAŞILAMADI"].shape[0]
        self.ozet.iloc[row, col] = ulasilamadi
        counts = self.tum_data['ALT KOD'].value_counts(dropna=True)
        counts_df = counts.reset_index()
        counts_df.columns = ['ALT KOD', 'Count']
        percent = (counts_df['Count'] / ulasilamadi) * 100
        counts_df['Percentage'] = percent
        counts_df['Percentage'] = counts_df['Percentage'].apply(lambda x: f"{x:.2f}%")
        counts_df.columns = [None, None, None]
        start_row = 2
        start_col_values = 4
        start_col_counts = 5
        start_col_percentage = 6
        for i in range(len(counts_df)):
            self.ozet.iloc[start_row + i, start_col_values] = counts_df.iloc[i, 0]
            self.ozet.iloc[start_row + i, start_col_counts] = counts_df.iloc[i, 1]
            self.ozet.iloc[start_row + i, start_col_percentage] = counts_df.iloc[i, 2]

    def make_pivot(self):
        # PIVOT: Sütun Etiketleri
        row, col = self.excel_to_index("A3")
        self.pivot.iloc[row, col] = "Sütun Etiketleri"
        self.pivot.iloc[row+1, col] = "Satın alma kanalı"
        self.pivot.iloc[row+2, col] = "example_product_1"
        self.pivot.iloc[row+3, col] = "example_product_2"
        self.pivot.iloc[row+4, col] = "another_product_example"
        self.pivot.iloc[row+5, col] = "Genel Toplam"
        row, col = self.excel_to_index("B5")
        example_product_1_count = self.basarili[self.basarili["Satın alma kanalı"] == "example_product_1"].shape[0]
        example_product_2_count = self.basarili[self.basarili["Satın alma kanalı"] == "example_product_2"].shape[0]
        other_count = self.basarili[~self.basarili["Satın alma kanalı"].isin(("example_product_2", "example_product_1"))].shape[0]
        self.pivot.iloc[row, col] = example_product_1_count
        self.pivot.iloc[row+1, col] = example_product_2_count
        self.pivot.iloc[row+2, col] = other_count
        self.pivot.iloc[row+3, col] = sum([example_product_1_count, example_product_2_count, other_count])
        self.pivot.iloc[self.excel_to_index("D1")] = "*ÜRÜN GRUBU"
        self.pivot.iloc[self.excel_to_index("I1")] = "*SATIŞ KANALI"
        self.pivot.iloc[self.excel_to_index("O1")] = "*MARKA"
        # PIVOT: *ÜRÜN GRUBU
        start_row, start_col = self.excel_to_index("D4")
        for unit in self.units:
            unique_values = self.basarili[unit].dropna().unique()
            self.unit_unique_values_urun_grubu[unit] = unique_values
            unit_df = pandas.DataFrame(columns=[unit, 'TV', 'BEYAZ ESYA', 'Genel Toplam'])
            for value in unique_values:
                tv_count = self.basarili[(self.basarili[unit] == value) & (self.basarili['Ürün grubu'] == 'TV')].shape[0]
                esya_count = self.basarili[(self.basarili[unit] == value) & (self.basarili['Ürün grubu'] == 'BEYAZ ESYA')].shape[0]
                total_count = tv_count + esya_count
                new_row = pandas.DataFrame({unit: [value], 'TV': [tv_count], 'BEYAZ ESYA': [esya_count], 'Genel Toplam': [total_count]})
                unit_df = pandas.concat([unit_df, new_row], ignore_index=True)
            total_tv = unit_df['TV'].sum()
            total_esya = unit_df['BEYAZ ESYA'].sum()
            total_all = total_tv + total_esya
            genel_toplam_row = pandas.DataFrame({unit: ['Genel Toplam'], 'TV': [total_tv], 'BEYAZ ESYA': [total_esya], 'Genel Toplam': [total_all]})
            unit_df = pandas.concat([unit_df, genel_toplam_row], ignore_index=True)
            self.pivot.iloc[start_row-1, start_col] = "Sütun Etiketleri"
            self.pivot.iloc[start_row-1, start_col+1] = "Ürün grubu"
            self.pivot.iloc[start_row, start_col] = unit
            self.pivot.iloc[start_row, start_col + 1] = 'TV'
            self.pivot.iloc[start_row, start_col + 2] = 'BEYAZ ESYA'
            self.pivot.iloc[start_row, start_col + 3] = 'Genel Toplam'
            start_row += 1
            for i in range(len(unit_df)):
                self.pivot.iloc[start_row + i, start_col] = unit_df.iloc[i, 0]  # Значение юнита
                self.pivot.iloc[start_row + i, start_col + 1] = unit_df.iloc[i, 1]  # TV
                self.pivot.iloc[start_row + i, start_col + 2] = unit_df.iloc[i, 2]  # BEYAZ ESYA
                self.pivot.iloc[start_row + i, start_col + 3] = unit_df.iloc[i, 3]  # Genel Toplam
            start_row += len(unit_df) + 3

        # PIVOT: *SATIŞ KANALI
        start_row, start_col = self.excel_to_index("I4")
        for unit in self.units:
            unique_values = self.basarili[unit].dropna().unique()
            self.unit_unique_values_satis_kanali[unit] = unique_values
            unit_df = pandas.DataFrame(columns=[unit, 'example_product_1', 'another_product_example', 'example_product_2', 'Genel Toplam'])
            for value in unique_values:
                example_product_1_count = self.basarili[(self.basarili[unit] == value) & (self.basarili['Satın alma kanalı'] == 'example_product_1')].shape[0]
                another_product_example_count = self.basarili[(self.basarili[unit] == value) & (~self.basarili['Satın alma kanalı'].isin(['example_product_1', 'example_product_2']))].shape[0]
                example_product_2_count = self.basarili[(self.basarili[unit] == value) & (self.basarili['Satın alma kanalı'] == 'example_product_2')].shape[0]
                total_count = example_product_1_count + another_product_example_count + example_product_2_count
                new_row = pandas.DataFrame({unit: [value], 'example_product_1': [example_product_1_count], 'another_product_example': [another_product_example_count], 'example_product_2': [example_product_2_count],'Genel Toplam': [total_count]})
                unit_df = pandas.concat([unit_df, new_row], ignore_index=True)
            total_example_product_1 = unit_df['example_product_1'].sum()
            total_another_product_example = unit_df['another_product_example'].sum()
            total_example_product_2 = unit_df['example_product_2'].sum()
            total_all = total_example_product_1 + total_another_product_example + total_example_product_2
            genel_toplam_row = pandas.DataFrame({unit: ['Genel Toplam'], 'example_product_1': [total_example_product_1], 'another_product_example': [total_another_product_example], 'example_product_2': [total_example_product_2], 'Genel Toplam': [total_all]})
            unit_df = pandas.concat([unit_df, genel_toplam_row], ignore_index=True)
            self.pivot.iloc[start_row-1, start_col] = "Sütun Etiketleri"
            self.pivot.iloc[start_row-1, start_col+1] = "Ürün grubu"
            self.pivot.iloc[start_row, start_col] = unit
            self.pivot.iloc[start_row, start_col + 1] = 'example_product_1'
            self.pivot.iloc[start_row, start_col + 2] = 'another_product_example'
            self.pivot.iloc[start_row, start_col + 3] = 'example_product_2'
            self.pivot.iloc[start_row, start_col + 4] = 'Genel Toplam'
            start_row += 1
            for i in range(len(unit_df)):
                self.pivot.iloc[start_row + i, start_col] = unit_df.iloc[i, 0]
                self.pivot.iloc[start_row + i, start_col + 1] = unit_df.iloc[i, 1]
                self.pivot.iloc[start_row + i, start_col + 2] = unit_df.iloc[i, 2]
                self.pivot.iloc[start_row + i, start_col + 3] = unit_df.iloc[i, 3]
                self.pivot.iloc[start_row + i, start_col + 4] = unit_df.iloc[i, 4]
            start_row += len(unit_df) + 3

        # PIVOT: *MARKA
        start_row, start_col = self.excel_to_index("O4")
        brands = self.basarili["MARKA"].dropna().unique()
        for unit in self.units:
            unique_values = self.basarili[unit].dropna().unique()
            columns = [unit] + list(brands) + ["Genel Toplam"]
            unit_df = pandas.DataFrame(columns=columns)
            for value in unique_values:
                row_data = {unit: value}
                total_count = 0
                for brand in brands:
                    count = self.basarili[(self.basarili[unit] == value) & (self.basarili["MARKA"] == brand)].shape[0]
                    row_data[brand] = count
                    total_count += count
                row_data["Genel Toplam"] = total_count
                unit_df = pandas.concat([unit_df, pandas.DataFrame([row_data])], ignore_index=True)
            genel_toplam_row = {unit: "Genel Toplam"}
            total_sum = 0
            for brand in brands:
                brand_count = unit_df[brand].sum()
                genel_toplam_row[brand] = brand_count
                total_sum += brand_count
            genel_toplam_row["Genel Toplam"] = total_sum
            unit_df = pandas.concat([unit_df, pandas.DataFrame([genel_toplam_row])])
            self.pivot.iloc[start_row-1, start_col] = "Sütun Etiketleri"
            self.pivot.iloc[start_row-1, start_col+1] = "Ürün grubu"
            self.pivot.iloc[start_row, start_col] = unit
            for i, brand in enumerate(brands):
                self.pivot.iloc[start_row, start_col + i + 1] = brand
            self.pivot.iloc[start_row, start_col + len(brands) + 1] = 'Genel Toplam'
            start_row += 1
            for row_idx in range(len(unit_df)):
                for col_idx, col_name in enumerate(unit_df.columns):
                    self.pivot.iloc[start_row + row_idx, start_col + col_idx] = unit_df.iloc[row_idx, col_idx]

            # Получаем строку с Genel Toplam, которая находится в последней строке
            genel_toplam = unit_df.loc[unit_df[unit] == "Genel Toplam"].iloc[0]

            # Создаем пустую таблицу для процентов
            percentage_df = unit_df.copy()

            # Проходимся по всем строкам, кроме строки Genel Toplam
            for row_idx in range(len(percentage_df) - 1):  # Последняя строка - Genel Toplam, её пропускаем
                for brand in brands:  # Проходимся по каждому бренду
                    current_value = unit_df.loc[row_idx, brand]
                    total_value = genel_toplam[brand]
                    if total_value > 0:
                        percentage_df.loc[row_idx, brand] = (current_value / total_value) * 100
                    else:
                        percentage_df.loc[row_idx, brand] = 0  # Если общий итог по бренду = 0, ставим 0

            # Убираем форматирование в процентном формате
            percentage_df = percentage_df.applymap(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)
            percentage_df_without_labels = percentage_df.iloc[:, 1:].copy()
            # Вставляем процентную таблицу в нужное место
            start_col_percent = start_col + len(brands) + 2  # Столбец, где начнется процентная таблица
            # Вставляем заголовки брендов для процентной таблицы
            for i, brand in enumerate(brands):
                self.pivot.iloc[start_row - 1, start_col_percent + i] = brand
            for row_idx in range(len(percentage_df_without_labels) - 1):  # Вставляем проценты, исключая строку Genel Toplam
                for col_idx, brand in enumerate(brands):
                    self.pivot.iloc[start_row + row_idx, start_col_percent + col_idx] = percentage_df_without_labels.iloc[row_idx, col_idx]
            start_row += len(unit_df) + 3

    def add_charts_to_excel(self, filepath):
        from openpyxl import load_workbook
        from openpyxl.drawing.image import Image
        import matplotlib.pyplot as plt
        wb = load_workbook(filepath)
        ws_grafikler = wb['GRAFIKLER']
        # добавление графика для Sütun Etiketleri - Satın alma kanalı
        start_row, start_col = self.excel_to_index("A5")
        example_product_1 = self.grafikler.iloc[start_row, start_col + 1]
        example_product_2 = self.grafikler.iloc[start_row + 1, start_col + 1]
        another_product_example = self.grafikler.iloc[start_row + 2, start_col + 1]
        labels = ['example_product_1', 'example_product_2', 'another_product_example']
        counts = [example_product_1, example_product_2, another_product_example]
        fig, ax = plt.subplots(figsize=(4, 3))
        ax.bar(labels, counts, color='blue')
        for i, v in enumerate(counts):
            ax.text(i, v + 1, str(v), ha='center', fontsize=8)
        ax.set_title('Satın Alma Kanali - Sütun Etiketleri', fontsize=10)
        chart_path = 'satin_alma_chart.png'
        plt.savefig(chart_path)
        plt.close(fig)
        img = Image(chart_path)
        ws_grafikler.add_image(img, 'N1')
        # добавление графиков Ürün grubu
        urun_grubu_cells = ["A17", "A30", "A46", "A60", "A73", "A87"]
        units_start_cell = dict(zip(self.units, urun_grubu_cells))
        start_row_grafik = 20
        for unit in self.units:
            start_row, start_col = self.excel_to_index(units_start_cell[unit])
            end_row = start_row + len(self.unit_unique_values_urun_grubu[unit]) + 1
            df = self.grafikler.iloc[start_row:end_row, start_col:start_col + 4]
            self.add_horizontal_chart_urun_grubu(ws_grafikler, df, f'{unit} - Ürün Grubu', f'N{start_row_grafik}')
            start_row_grafik += 20
        # добавление графиков satis kanali
        satis_kanali_cells = ["A102", "A117", "A133", "A149", "A162", "A177"]
        units_start_cell = dict(zip(self.units, satis_kanali_cells))
        start_row_grafik = 150
        for unit in self.units:
            start_row, start_col = self.excel_to_index(units_start_cell[unit])
            end_row = start_row + len(self.unit_unique_values_urun_grubu[unit]) + 1
            df = self.grafikler.iloc[start_row:end_row, start_col:start_col + 5]
            self.add_horizontal_chart_satis_kanali(ws_grafikler, df, f'{unit} - SATIŞ KANALI', f'N{start_row_grafik}')
            start_row_grafik += 40

        wb.save(filepath)

    def add_horizontal_chart_urun_grubu(self, ws, df, title, start_cell):
        import matplotlib.pyplot as plt
        from openpyxl.drawing.image import Image
        import numpy as np
        import pandas as pd  # Не забудьте импортировать pandas

        # Исключаем строку с "Genel Toplam"
        filtered_df = df[df.iloc[:, 0] != 'Genel Toplam']

        # Данные для графика
        labels = filtered_df.iloc[:, 0]  # Ответы
        tv_counts = filtered_df.iloc[:, 1]  # Количество TV
        esya_counts = filtered_df.iloc[:, 2]  # Количество BEYAZ ESYA

        # Удаляем NaN значения и создаем DataFrame для совместимости
        combined_df = pd.DataFrame({
            'labels': labels,
            'tv_counts': tv_counts,
            'esya_counts': esya_counts
        })

        # Заменяем строки с пробелами на NaN и удаляем пустые строки
        combined_df.replace(' ', np.nan, inplace=True)
        combined_df.dropna(inplace=True)

        # Приводим данные к числовому типу
        combined_df['tv_counts'] = pd.to_numeric(combined_df['tv_counts'], errors='coerce')
        combined_df['esya_counts'] = pd.to_numeric(combined_df['esya_counts'], errors='coerce')

        # Снова удаляем NaN значения после преобразования типов
        combined_df.dropna(inplace=True)

        labels = combined_df['labels']
        tv_counts = combined_df['tv_counts']
        esya_counts = combined_df['esya_counts']

        # Создаём сдвиги для осей
        y_pos = np.arange(len(labels))  # Позиции для TV
        y_pos_esya = y_pos - 0.5  # Позиции для BEYAZ ESYA, с небольшим сдвигом

        # Создаем график
        fig, ax = plt.subplots(figsize=(20, 3))  # Размер графика

        # Построение двойных столбцов с небольшим сдвигом
        ax.barh(y_pos, tv_counts, height=0.4, color='blue', label='TV')
        ax.barh(y_pos_esya, esya_counts, height=0.4, color='red', label='BEYAZ ESYA')

        # Добавляем подписи значений
        for i, (tv, esya) in enumerate(zip(tv_counts, esya_counts)):
            ax.text(tv + 5, y_pos[i], str(tv), va='center', fontsize=8, color='blue')
            ax.text(esya + 5, y_pos_esya[i], str(esya), va='center', fontsize=8, color='red')

        # Уменьшение шрифта для меток оси Y
        ax.set_yticks(y_pos - 0.1)  # Центрируем метки между столбцами
        ax.set_yticklabels(labels, fontsize=8)

        # Заголовок
        ax.set_title(title)

        # Добавляем легенду для обозначения цветов
        ax.legend(loc='upper right')

        # Сохранение графика
        chart_path = f'{title}.png'
        plt.savefig(chart_path)
        plt.close(fig)

        # Вставка графика в Excel
        img = Image(chart_path)
        ws.add_image(img, start_cell)

    def add_horizontal_chart_satis_kanali(self, ws, df, title, start_cell):
        import matplotlib.pyplot as plt
        from openpyxl.drawing.image import Image
        import numpy as np
        
        # Исключаем строку с "Genel Toplam"
        filtered_df = df[df.iloc[:, 0] != 'Genel Toplam']
        
        # Данные для графика
        labels = filtered_df.iloc[:, 0]  # Ответы
        example_product_1_counts = filtered_df.iloc[:, 1]  # Количество example_product_1
        another_product_example_counts = filtered_df.iloc[:, 2]  # Количество another_product_example
        example_product_2_counts = filtered_df.iloc[:, 3]  # Количество example_product_2

        combined_df = pandas.DataFrame({
            'labels': labels,
            'example_product_1_counts': example_product_1_counts,
            'another_product_example_counts': another_product_example_counts,
            'example_product_2_counts': example_product_2_counts
        })
        combined_df.replace(' ', np.nan, inplace=True)
        combined_df.dropna(inplace=True)
        combined_df['example_product_1_counts'] = pandas.to_numeric(combined_df['example_product_1_counts'], errors='coerce')
        combined_df['another_product_example_counts'] = pandas.to_numeric(combined_df['another_product_example_counts'], errors='coerce')
        combined_df['example_product_2_counts'] = pandas.to_numeric(combined_df['example_product_2_counts'], errors='coerce')
        combined_df.dropna(inplace=True)
        labels = combined_df['labels']
        example_product_1_counts = combined_df['example_product_1_counts']
        another_product_example_counts = combined_df['another_product_example_counts']
        example_product_2_counts = combined_df['example_product_2_counts']
        # Создаём сдвиги для осей
        y_pos = np.arange(len(labels))  # Позиции для example_product_1
        y_pos_another_product_example = y_pos - 0.3  # Позиции для another_product_example
        y_pos_example_product_2 = y_pos_another_product_example - 0.3  # Позиции для another_product_example

        # Создаем график
        fig, ax = plt.subplots(figsize=(25, 7))  # Размер графика

        # Построение тройных столбцов с небольшим сдвигом
        ax.barh(y_pos, example_product_1_counts, height=0.4, color='blue', label='example_product_1')
        ax.barh(y_pos_another_product_example, another_product_example_counts, height=0.4, color='red', label='another_product_example')
        ax.barh(y_pos_example_product_2, example_product_2_counts, height=0.4, color='green', label='example_product_2')

        # Добавляем подписи значений
        for i, (example_product_1, another_product_example, example_product_2) in enumerate(zip(example_product_1_counts, another_product_example_counts, example_product_2_counts)):
            ax.text(example_product_1 + 5, y_pos[i], str(example_product_1), va='center', fontsize=8, color='blue')
            ax.text(another_product_example + 5, y_pos_another_product_example[i], str(another_product_example), va='center', fontsize=8, color='red')
            ax.text(example_product_2 + 5, y_pos_example_product_2[i], str(example_product_2), va='center', fontsize=8, color='green')

        # Уменьшение шрифта для меток оси Y
        ax.set_yticks(y_pos - 0.1)  # Центрируем метки между столбцами
        ax.set_yticklabels(labels, fontsize=8)

        # Заголовок
        ax.set_title(title)

        # Добавляем легенду для обозначения цветов
        ax.legend(loc='upper right')

        # Сохранение графика
        chart_path = f'{title}.png'
        plt.savefig(chart_path)
        plt.close(fig)

        # Вставка графика в Excel
        img = Image(chart_path)
        ws.add_image(img, start_cell)

    def make_grafikler(self):

        # GRAFIKLER: Sütun Etiketleri
        row, col = self.excel_to_index("A3")
        self.grafikler.iloc[row, col] = "Sütun Etiketleri"
        self.grafikler.iloc[row+1, col] = "Satın alma kanalı"
        self.grafikler.iloc[row+2, col] = "example_product_1"
        self.grafikler.iloc[row+3, col] = "example_product_2"
        self.grafikler.iloc[row+4, col] = "another_product_example"
        self.grafikler.iloc[row+5, col] = "Genel Toplam"
        row, col = self.excel_to_index("B5")
        example_product_1_count = self.basarili[self.basarili["Satın alma kanalı"] == "example_product_1"].shape[0]
        example_product_2_count = self.basarili[self.basarili["Satın alma kanalı"] == "example_product_2"].shape[0]
        other_count = self.basarili[~self.basarili["Satın alma kanalı"].isin(("example_product_2", "example_product_1"))].shape[0]
        genel_toplam = sum([example_product_1_count, example_product_2_count, other_count])
        example_product_1_percent = f"{(example_product_1_count / genel_toplam * 100):.2f}%"
        example_product_2_percent = f"{(example_product_2_count / genel_toplam * 100):.2f}%"
        other_percent = f"{(other_count / genel_toplam * 100):.2f}%"
        self.grafikler.iloc[row, col] = example_product_1_count
        self.grafikler.iloc[row, col+1] = example_product_1_percent
        self.grafikler.iloc[row+1, col] = example_product_2_count
        self.grafikler.iloc[row+1, col+1] = example_product_2_percent
        self.grafikler.iloc[row+2, col] = other_count
        self.grafikler.iloc[row+2, col+1] = other_percent
        self.grafikler.iloc[row+3, col] = genel_toplam

        # GRAFIKLER: *ÜRÜN GRUBU
        grafikler_row = 15
        grafikler_col = 0
        start_row, start_col = self.excel_to_index("A2")
        for unit in self.units:
            unique_values = self.basarili[unit].dropna().unique()
            unit_df = pandas.DataFrame(columns=[unit, 'TV', 'BEYAZ ESYA', 'Genel Toplam'])
            for value in unique_values:
                tv_count = self.basarili[(self.basarili[unit] == value) & (self.basarili['Ürün grubu'] == 'TV')].shape[0]
                esya_count = self.basarili[(self.basarili[unit] == value) & (self.basarili['Ürün grubu'] == 'BEYAZ ESYA')].shape[0]
                total_count = tv_count + esya_count
                new_row = pandas.DataFrame({unit: [value], 'TV': [tv_count], 'BEYAZ ESYA': [esya_count], 'Genel Toplam': [total_count]})
                unit_df = pandas.concat([unit_df, new_row], ignore_index=True)
            total_tv = unit_df['TV'].sum()
            total_esya = unit_df['BEYAZ ESYA'].sum()
            total_all = total_tv + total_esya
            genel_toplam_row = pandas.DataFrame({unit: ['Genel Toplam'], 'TV': [total_tv], 'BEYAZ ESYA': [total_esya], 'Genel Toplam': [total_all]})
            unit_df = pandas.concat([unit_df, genel_toplam_row], ignore_index=True)
            # Добавляем таблицы в self.grafikler
            self.grafikler.iloc[grafikler_row-1, start_col] = "Sütun Etiketleri"
            self.grafikler.iloc[grafikler_row-1, start_col+1] = "Ürün grubu"
            self.grafikler.iloc[grafikler_row, grafikler_col] = unit
            self.grafikler.iloc[grafikler_row, grafikler_col+1] = "TV"
            self.grafikler.iloc[grafikler_row, grafikler_col+2] = "BEYAZ ESYA"
            self.grafikler.iloc[grafikler_row, grafikler_col+3] = "Genel Toplam"
            grafikler_row += 1
            
            # Перенос данных
            for i in range(len(unit_df)):
                self.grafikler.iloc[grafikler_row + i, grafikler_col] = unit_df.iloc[i, 0]
                self.grafikler.iloc[grafikler_row + i, grafikler_col + 1] = unit_df.iloc[i, 1]
                self.grafikler.iloc[grafikler_row + i, grafikler_col + 2] = unit_df.iloc[i, 2]
                self.grafikler.iloc[grafikler_row + i, grafikler_col + 3] = unit_df.iloc[i, 3]
            # Теперь добавляем расчет процентов рядом с основной таблицей
            genel_toplam = unit_df.loc[unit_df[unit] == 'Genel Toplam'].iloc[0]  # Строка с итогами

            percentage_df = unit_df.copy().iloc[:-1]  # Исключаем строку "Genel Toplam"
            for idx, row in percentage_df.iterrows():
                for col in row.index[1:]:  # Проходим по каждому столбцу, кроме первого (названия юнита)
                    current_value = row[col]
                    total_value = genel_toplam[col]
                    percentage_df.at[idx, col] = (current_value / total_value) * 100 if total_value > 0 else 0

            # Записываем процентные данные в self.grafikler
            start_col_percent = grafikler_col + 4  # Стартовый столбец для процентов
            self.grafikler.iloc[grafikler_row - 1, start_col_percent:start_col_percent + 3] = ["TV", "BEYAZ ESYA", "Genel Toplam"]

            for i in range(len(percentage_df)):
                self.grafikler.iloc[grafikler_row + i, start_col_percent] = f"{percentage_df.iloc[i, 1]:.2f}%"
                self.grafikler.iloc[grafikler_row + i, start_col_percent + 1] = f"{percentage_df.iloc[i, 2]:.2f}%"
                self.grafikler.iloc[grafikler_row + i, start_col_percent + 2] = f"{percentage_df.iloc[i, 3]:.2f}%"

            # Добавляем отступ между таблицами
            grafikler_row += len(unit_df) + 5

        # GRAFIKLER: *SATIŞ KANALI
        start_row, start_col = 100, 0
        for unit in self.units:
            unique_values = self.basarili[unit].dropna().unique()
            unit_df = pandas.DataFrame(columns=[unit, 'example_product_1', 'another_product_example', 'example_product_2', 'Genel Toplam'])
            for value in unique_values:
                example_product_1_count = self.basarili[(self.basarili[unit] == value) & (self.basarili['Satın alma kanalı'] == 'example_product_1')].shape[0]
                another_product_example_count = self.basarili[(self.basarili[unit] == value) & (~self.basarili['Satın alma kanalı'].isin(['example_product_1', 'example_product_2']))].shape[0]
                example_product_2_count = self.basarili[(self.basarili[unit] == value) & (self.basarili['Satın alma kanalı'] == 'example_product_2')].shape[0]
                total_count = example_product_1_count + another_product_example_count + example_product_2_count
                new_row = pandas.DataFrame({unit: [value], 'example_product_1': [example_product_1_count], 'another_product_example': [another_product_example_count], 'example_product_2': [example_product_2_count],'Genel Toplam': [total_count]})
                unit_df = pandas.concat([unit_df, new_row], ignore_index=True)
            total_example_product_1 = unit_df['example_product_1'].sum()
            total_another_product_example = unit_df['another_product_example'].sum()
            total_example_product_2 = unit_df['example_product_2'].sum()
            total_all = total_example_product_1 + total_another_product_example + total_example_product_2
            genel_toplam_row = pandas.DataFrame({unit: ['Genel Toplam'], 'example_product_1': [total_example_product_1], 'another_product_example': [total_another_product_example], 'example_product_2': [total_example_product_2], 'Genel Toplam': [total_all]})
            unit_df = pandas.concat([unit_df, genel_toplam_row], ignore_index=True)
            self.grafikler.iloc[start_row-1, start_col] = "Sütun Etiketleri"
            self.grafikler.iloc[start_row-1, start_col+1] = "Ürün grubu"
            self.grafikler.iloc[start_row, start_col] = unit
            self.grafikler.iloc[start_row, start_col + 1] = 'example_product_1'
            self.grafikler.iloc[start_row, start_col + 2] = 'another_product_example'
            self.grafikler.iloc[start_row, start_col + 3] = 'example_product_2'
            self.grafikler.iloc[start_row, start_col + 4] = 'Genel Toplam'
            start_row += 1
            for i in range(len(unit_df)):
                self.grafikler.iloc[start_row + i, start_col] = unit_df.iloc[i, 0]
                self.grafikler.iloc[start_row + i, start_col + 1] = unit_df.iloc[i, 1]
                self.grafikler.iloc[start_row + i, start_col + 2] = unit_df.iloc[i, 2]
                self.grafikler.iloc[start_row + i, start_col + 3] = unit_df.iloc[i, 3]
                self.grafikler.iloc[start_row + i, start_col + 4] = unit_df.iloc[i, 4]
            # Добавляем расчет процентов
            genel_toplam = unit_df.loc[unit_df[unit] == 'Genel Toplam'].iloc[0]  # Строка с итогами

            percentage_df = unit_df.copy().iloc[:-1]  # Исключаем строку "Genel Toplam"
            for idx, row in percentage_df.iterrows():
                for col in row.index[1:]:
                    current_value = row[col]
                    total_value = genel_toplam[col]
                    percentage_df.at[idx, col] = (current_value / total_value) * 100 if total_value > 0 else 0

            # Записываем процентные данные
            start_col_percent = start_col + 5
            self.grafikler.iloc[start_row - 1, start_col_percent:start_col_percent + 4] = ["example_product_1", "another_product_example", "example_product_2", "Genel Toplam"]

            for i in range(len(percentage_df)):
                self.grafikler.iloc[start_row + i, start_col_percent] = f"{percentage_df.iloc[i, 1]:.2f}%"
                self.grafikler.iloc[start_row + i, start_col_percent + 1] = f"{percentage_df.iloc[i, 2]:.2f}%"
                self.grafikler.iloc[start_row + i, start_col_percent + 2] = f"{percentage_df.iloc[i, 3]:.2f}%"
                self.grafikler.iloc[start_row + i, start_col_percent + 3] = f"{percentage_df.iloc[i, 4]:.2f}%"


            start_row += len(unit_df) + 6

    def make_result(self):
        self.make_tum_data()
        self.make_tum_data_tekil()
        self.make_ulasilan_data()
        self.make_basarili()
        self.make_ozet()
        self.make_pivot()
        self.make_grafikler()
        with pandas.ExcelWriter('output.xlsx') as writer:
            self.tum_data.to_excel(writer, sheet_name='Tüm Data', index=False)
            self.tum_data_tekil.to_excel(writer, sheet_name='TÜM DATA TEKİL', index=False)
            self.ulasilan_data.to_excel(writer, sheet_name='ULAŞILAN DATA', index=False)
            self.basarili.to_excel(writer, sheet_name='BAŞARILI', index=False)
            self.ozet.to_excel(writer, sheet_name='ÖZET', index=False, header=False)
            self.pivot.to_excel(writer, sheet_name='Pivot', index=False, header=False)
            self.grafikler.to_excel(writer, sheet_name='GRAFIKLER', index=False, header=False)
        self.add_charts_to_excel('output.xlsx')


def start():
    LOGGER.info("start")
    refactor = Main("origin_analytics.xlsx")
    refactor.make_result()

start()
