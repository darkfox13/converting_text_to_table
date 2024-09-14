from odf import draw, text, table
from odf.draw import Page
from odf.opendocument import OpenDocumentText, OpenDocumentPresentation
from odf.style import Style, TextProperties
from odf.text import P
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell, TableColumn

doc = load("spells test.odt")

# Extract paragraphs from the document
all_text = []
for element in doc.getElementsByType(P):
    all_text.append(str(element))

# Join all paragraphs into one string
joined_text = '\n'.join(all_text)
#sperating all the spells
list_of_spell = joined_text.split("##")
#creating a dict with number for keys
spell_dict = {i: item for i, item in enumerate(list_of_spell)}
#spliting the spells into their sperate parts in the dict
for key in spell_dict:
    spell_dict[key] = (spell_dict[key]).rsplit("\n")

#print(spell_dict)
print(spell_dict[0])
print(spell_dict[1])
print(spell_dict[2])
print(len(spell_dict))

###############################

doc = OpenDocumentText()

table_style = Style(name="TableStyle", family="table")
doc.styles.addElement(table_style)


def build_table(i):
    # Create a page break style with the correct fo:break-before property
    page_break_style = Style(name="PageBreak", family="paragraph")
    doc.styles.addElement(page_break_style)

    # Add some content before the page break
    paragraph1 = P(text="")
    doc.text.addElement(paragraph1)

    # Create a table and apply the style
    table = Table(name="ExampleTable", stylename=table_style)

    # Define the number of columns
    for _ in range(3):
        table.addElement(TableColumn())

    #  first row of the table
    row1 = TableRow()
    row1_data = [spell_dict[i][3]]
    for cell_data in row1_data:
        cell = TableCell()
        cell.addElement(P(text=cell_data))
        row1.addElement(cell)
    table.addElement(row1)

    merged_cell = TableCell(numbercolumnsspanned=3)  # Merge 3 columns
    merged_cell.addElement(P(text=spell_dict[i][1]))
    row1.addElement(merged_cell)
    table.addElement(row1)

    # second row of the table

    row2 = TableRow()
    row2_data = [""]
    for cell_data in row2_data:
        cell = TableCell()
        cell.addElement(P(text=spell_dict[i][2]))
        row2.addElement(cell)
    table.addElement(row2)

    merged_cell = TableCell(numbercolumnsspanned=3)  # Merge 3 columns
    merged_cell.addElement(P(text=spell_dict[i][6]))
    row2.addElement(merged_cell)
    table.addElement(row2)

    # third row of the table

    row3 = TableRow()
    row3_data = [spell_dict[i][4], spell_dict[i][7], spell_dict[i][5]]
    for cell_data in row3_data:
        cell = TableCell()
        cell.addElement(P(text=cell_data))
        row3.addElement(cell)
    table.addElement(row3)

    # fourth row of the table
    row4 = TableRow()
    row4_data = []
    for cell_data in row4_data:
        cell = TableCell()
        cell.addElement(P(text=cell_data))
        row4.addElement(cell)
    table.addElement(row4)

    merged_cell = TableCell(numbercolumnsspanned=4)  # Merge 4 columns
    merged_cell.addElement(P(text=spell_dict[i][8]))
    row4.addElement(merged_cell)
    table.addElement(row4)

    # fifth row

    row5 = TableRow()
    row5_data = []
    for cell_data in row5_data:
        cell = TableCell()
        cell.addElement(P(text=cell_data))
        row4.addElement(cell)
    table.addElement(row5)

    merged_cell = TableCell(numbercolumnsspanned=4)  # Merge 4 columns
    merged_cell.addElement(P(text=spell_dict[i][9]))
    row5.addElement(merged_cell)
    table.addElement(row5)

    # Add the table to the document
    doc.text.addElement(table)

    page_break_paragraph = P(text="This content will start on a new page.", stylename=page_break_style)
    doc.text.addElement(page_break_paragraph)

    paragraph2 = P(text="")
    doc.text.addElement(paragraph2)



for key in range(1, len(spell_dict)):
    build_table(key)

# Save the document
doc.save("test.odt")

print("Table saved as test.odt")
