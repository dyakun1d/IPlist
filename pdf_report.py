from fpdf import FPDF

def create_pdf(author_data, main_table_data, page_format="A3"):
    # Создаем PDF объект с форматом A3
    pdf = FPDF(orientation='L', format=page_format)
    pdf.add_page()

    # Добавляем шрифт DejaVuSans для поддержки Unicode
    pdf.add_font("DejaVu", "", "DejaVuSans.ttf", uni=True)
    pdf.set_font("DejaVu", size=10)

    # Добавляем малую таблицу с данными автора и мощностью проекта
    pdf.cell(0, 10, txt=f"Author: {author_data['author']}", ln=True)
    pdf.cell(0, 10, txt=f"Creation Date: {author_data['creation_date']}", ln=True)
    pdf.cell(0, 10, txt=f"Project Name: {author_data['project_name']}", ln=True)
    pdf.cell(0, 10, txt=f"Power (Wp): {author_data['power_wp']}", ln=True)
    pdf.ln(20)

    # Добавляем основную таблицу
    col_width = pdf.w / len(main_table_data[0]) - 1  # Вычисляем ширину столбца
    row_height = pdf.font_size + 2
    for row in main_table_data:
        for item in row:
            # Приводим текст к строке
            text = str(item)
            pdf.cell(col_width, row_height, txt=text, border=1)
        pdf.ln(row_height)

    # Возвращаем PDF как байтовый объект
    pdf_output = pdf.output(dest='S')
    return pdf_output
