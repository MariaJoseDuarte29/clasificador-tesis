import sys
import pandas as pd

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='utf-8'>
<title>Excel Table</title>
<link rel='stylesheet' type='text/css' href='https://cdn.datatables.net/1.13.6/css/jquery.dataTables.css'>
<style>
body { font-family: Arial, sans-serif; margin: 40px; }
</style>
</head>
<body>
<table id='table' class='display'>
{table}
</table>
<script src='https://code.jquery.com/jquery-3.7.0.min.js'></script>
<script src='https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js'></script>
<script>
$(document).ready(function() {
  $('#table').DataTable({
    scrollX: true
  });
});
</script>
</body>
</html>
"""

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('Uso: python xlsx_to_html.py archivo.xlsx salida.html')
        sys.exit(1)

    xlsx_path, output_html = sys.argv[1], sys.argv[2]
    df = pd.read_excel(xlsx_path)
    html_table = df.to_html(index=False)

    with open(output_html, 'w', encoding='utf-8') as f:
        f.write(HTML_TEMPLATE.format(table=html_table))

    print(f'Archivo {output_html} generado.')
