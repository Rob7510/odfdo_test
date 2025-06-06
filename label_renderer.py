from odfdo import Document, Table, Row, Cell, Paragraph, Style


def mm_to_cm(mm):
    return round(mm / 10, 2)

def generate_label_table(layout):
    doc = Document('blank_A4.odt')
    body = doc.body

    table = Table(name=layout['name'], width=layout['t_columns'], height=layout['t_rows'])
    print('Table size', table.size)

    # Define and apply column styles
    col_width_cm = mm_to_cm(layout['label_width_mm'] + layout['horizontal_spacing_mm'])
    for i in range(layout['t_columns']):
        col_style = Style(family="table-column", name=f"ColWidth{i}")
        col_style.set_properties({"style:column-width": f"{col_width_cm}cm"}, area="table-column")
        doc.insert_style(col_style, automatic=True)
        col = table.get_column(i)
        col.style = col_style.name

    # Set fixed row height and fill pre-initialized cells
    row_height_cm = f"{mm_to_cm(layout['label_height_mm'] + layout['vertical_spacing_mm'])}cm"
    for y in range(layout['t_rows']):
        row = table.get_row(y)
        row.set_style_attribute("min-row-height", row_height_cm)
        for x in range(layout['t_columns']):
            cell = row.get_cell(x)
            para = Paragraph("A")
            if layout['outline']:
                cell.set_style_attribute("border", "0.06pt solid #000000")
            cell.append(para)

    table.set_style_attribute("margin-left", f"{mm_to_cm(layout['left_margin_mm'])}cm")
    table.set_style_attribute("margin-top", f"{mm_to_cm(layout['top_margin_mm'])}cm")
    table.set_style_attribute("width", f"{mm_to_cm(layout['t_columns'] * layout['label_width_mm'] + (layout['t_columns'] - 1) * layout.get('horizontal_spacing_mm', 0))}cm")

    body.append(table)

    print('debugging')
    tables = body.get_tables()
    tablenames = [t.name for t in tables]
    print('Tablenames', tablenames)

    output_file = f"label_form_{layout['name']}.odt"
    doc.save(output_file)
    print(f"Saved: {output_file}")

if __name__ == '__main__':
    my_def = {'name':'test2',
      't_rows':2,
      't_columns':3,
      'label_width_mm':30,
      'label_height_mm':18,
      'left_margin_mm':6,
      'top_margin_mm':2,
      'horizontal_spacing_mm':2,
      'vertical_spacing_mm':3,
      'outline':1
      }
    print('Starting')
    generate_label_table(my_def)