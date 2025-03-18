using OfficeOpenXml;

namespace procesadorDeExel
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private string path = "";
        private List<object[]> DesdoblarCorreos(ExcelWorksheet hoja)
        {
            var registrosProcesados = new List<object[]>();

            for (int fila = 2; fila <= hoja.Dimension.End.Row; fila++)
            {
                var codigo = hoja.Cells[fila, 1].Text;
                var fBaja = hoja.Cells[fila, 2].Text;
                var plan = hoja.Cells[fila, 3].Text;
                var descripcion = hoja.Cells[fila, 4].Text;
                var emails = hoja.Cells[fila, 5].Text;

                var listaCorreos = emails.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                         .Select(e => e.Trim())
                                         .ToArray();

                foreach (var mail in listaCorreos)
                {
                    registrosProcesados.Add(new object[] { codigo, fBaja, plan, descripcion, mail });
                }
            }

            return registrosProcesados;
        }

        private List<object[]> ActualizarPlanSiHayBaja(List<object[]> registros)
        {
            foreach (var registro in registros)
            {
                var fBaja = registro[1]?.ToString();

                if (DateTime.TryParse(fBaja, out _))
                {
                    registro[2] = "baja";
                }
            }

            return registros;
        }
        private void btnCargarExcelNuevo_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fileDialog = new OpenFileDialog())
            {
                fileDialog.Filter = "Archivos Excel (*.xlsx)|*.xlsx";
                fileDialog.Title = "Seleccionar un archivo excel";
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtExcelNuevo.Text = fileDialog.SafeFileName;
                    path = fileDialog.FileName;
                }
            }
        }

        private void ProcesarExcel(string rutaExcel, string rutaOutput)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(rutaExcel)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    MessageBox.Show("El archivo Excel no contiene hojas.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                var hoja = package.Workbook.Worksheets[0];

                var registrosDesdoblados = DesdoblarCorreos(hoja);

                var registrosFinales = ActualizarPlanSiHayBaja(registrosDesdoblados);

                using (var nuevoPackage = new ExcelPackage())
                {
                    var nuevaHoja = nuevoPackage.Workbook.Worksheets.Add("Modificado");

                    nuevaHoja.Cells[1, 1].Value = "Codigo";
                    nuevaHoja.Cells[1, 2].Value = "FBaja";
                    nuevaHoja.Cells[1, 3].Value = "Plan";
                    nuevaHoja.Cells[1, 4].Value = "Descripcion";
                    nuevaHoja.Cells[1, 5].Value = "Mail";

                    for (int i = 0; i < registrosFinales.Count; i++)
                    {
                        for (int j = 0; j < registrosFinales[i].Length; j++)
                        {
                            nuevaHoja.Cells[i + 2, j + 1].Value = registrosFinales[i][j];
                        }
                    }

                    nuevoPackage.SaveAs(new FileInfo(rutaOutput));
                }
            }
        }

        private void btnProcesar_Click(object sender, EventArgs e)
        {
            string rutaExcel = path;

            if (string.IsNullOrEmpty(rutaExcel))
            {
                MessageBox.Show("Por favor, selecciona un archivo Excel antes de continuar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                string rutaOutput = Path.Combine(Path.GetDirectoryName(rutaExcel), $"{txtExcelNuevo.Text}_Modificado.xlsx");
                ProcesarExcel(rutaExcel, rutaOutput);

                MessageBox.Show($"El Excel modificado se ha generado con éxito:\n{rutaOutput}", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error durante el procesamiento: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private List<object[]> LeerRegistrosDesdeHoja(ExcelWorksheet hoja)
        {
            var registros = new List<object[]>();

            for (int i = 2; i <= hoja.Dimension.End.Row; i++)
            {
                var fila = new object[hoja.Dimension.End.Column];
                for (int j = 1; j <= hoja.Dimension.End.Column; j++)
                {
                    fila[j - 1] = hoja.Cells[i, j].Value;
                }
                registros.Add(fila);
            }

            return registros;
        }
        private List<object[]> GenerarArchivoResultado(List<object[]> registrosViejo, List<object[]> registrosNuevo)
        {
            var registrosResultado = new List<object[]>(registrosNuevo); // Clonar registros nuevos
            var dictMailsViejos = new Dictionary<string, string>(); // Documento -> Mail viejo

            // Crear un diccionario de documentos y mails del Excel viejo
            foreach (var registro in registrosViejo)
            {
                string codigo = registro[0]?.ToString(); // Asumiendo columna 0 es "codigo"
                string mail = registro[4]?.ToString(); // Asumiendo columna 3 es "mail"
                if (!string.IsNullOrEmpty(codigo) && !string.IsNullOrEmpty(mail))
                {
                    dictMailsViejos[codigo] = mail;
                }
            }

            // Buscar diferencias entre el mail del Excel nuevo y viejo
            foreach (var registroNuevo in registrosNuevo)
            {
                string codigo = registroNuevo[0]?.ToString(); // Asumiendo columna 0 es "codigo"
                string mailNuevo = registroNuevo[4]?.ToString(); // Asumiendo columna 3 es "mail"

                if (dictMailsViejos.ContainsKey(codigo) && dictMailsViejos[codigo] != mailNuevo)
                {
                    string mailViejo = dictMailsViejos[codigo];
                    var nuevoRegistro = (object[])registroNuevo.Clone(); // Clonar registro actual
                    nuevoRegistro[4] = mailViejo; // Cambiar el mail al del Excel viejo
                    nuevoRegistro[2] = "Baja"; // Cambiar el plan a "Baja" (columna 1, ajusta si es diferente)
                    registrosResultado.Add(nuevoRegistro);
                }
            }

            return registrosResultado;
        }
        private void ProcesarDosExcels(string rutaViejo, string rutaNuevo, string rutaOutput)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var packageViejo = new ExcelPackage(new FileInfo(rutaViejo)))
            using (var packageNuevo = new ExcelPackage(new FileInfo(rutaNuevo)))
            {
                var hojaViejo = packageViejo.Workbook.Worksheets[0];
                var hojaNuevo = packageNuevo.Workbook.Worksheets[0];

                // Leer registros de ambos Excels
                var registrosViejo = LeerRegistrosDesdeHoja(hojaViejo);
                var registrosNuevo = LeerRegistrosDesdeHoja(hojaNuevo);

                // Generar registros modificados
                var registrosResultado = GenerarArchivoResultado(registrosViejo, registrosNuevo);

                // Crear nuevo archivo Excel
                using (var nuevoPackage = new ExcelPackage())
                {
                    var nuevaHoja = nuevoPackage.Workbook.Worksheets.Add("Resultado");

                    // Escribir encabezados
                    nuevaHoja.Cells[1, 1].Value = "Codigo";
                    nuevaHoja.Cells[1, 2].Value = "FBaja";
                    nuevaHoja.Cells[1, 3].Value = "Plan";
                    nuevaHoja.Cells[1, 4].Value = "Descripcion";
                    nuevaHoja.Cells[1, 5].Value = "Mail";

                    // Escribir registros procesados
                    for (int i = 0; i < registrosResultado.Count; i++)
                    {
                        for (int j = 0; j < registrosResultado[i].Length; j++)
                        {
                            nuevaHoja.Cells[i + 2, j + 1].Value = registrosResultado[i][j];
                        }
                    }

                    // Guardar el archivo generado
                    nuevoPackage.SaveAs(new FileInfo(rutaOutput));
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Archivos Excel (*.xlsx)|*.xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string rutaViejo = textBox1.Text;
            string rutaNuevo = textBox2.Text;

            if (string.IsNullOrEmpty(rutaViejo) || string.IsNullOrEmpty(rutaNuevo))
            {
                MessageBox.Show("Por favor, selecciona ambos archivos Excel.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                string rutaOutput = Path.Combine(Path.GetDirectoryName(rutaNuevo), "ExcelModificado.xlsx");
                ProcesarDosExcels(rutaViejo, rutaNuevo, rutaOutput);

                MessageBox.Show($"El archivo modificado se ha generado con éxito:\n{rutaOutput}", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocurrió un error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Archivos Excel (*.xlsx)|*.xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = ofd.FileName;
            }
        }
    }
}