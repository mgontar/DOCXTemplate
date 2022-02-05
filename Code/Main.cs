using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;

namespace DOCXTemplate
{
    public partial class Main : Form
    {

        BindingList<TemplateVariable> templateVariableNames;
        String templateDirectoryPath;
        String outputDirectoryPath;

        public Main()
        {
            InitializeComponent();

            templateDirectoryPath = AppDomain.CurrentDomain.BaseDirectory + @"\templates";

            if (!Directory.Exists(templateDirectoryPath))
            {
                Directory.CreateDirectory(templateDirectoryPath);
            }

            outputDirectoryPath = AppDomain.CurrentDomain.BaseDirectory + @"\output";
            if (!Directory.Exists(outputDirectoryPath))
            {
                Directory.CreateDirectory(outputDirectoryPath);
            }

            this.Load += new EventHandler(Main_Load);
        }

        private void Main_Load(System.Object sender, System.EventArgs e)
        {
            scanTemplates();
        }

        private void btnScanTemplates_Click(object sender, EventArgs e)
        {
            scanTemplates();
        }

        private void clbTemplates_SelectedIndexChanged(object sender, EventArgs e)
        {
            refreshVariablesForAllTemplates();
        }

        private void btnToggleSelect_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < clbTemplates.Items.Count; i++) {
                var newChecked = !(clbTemplates.GetItemCheckState(i) == CheckState.Checked);
                clbTemplates.SetItemChecked(i, newChecked);
            }
            refreshVariablesForAllTemplates();
        }

        private void refreshVariablesForAllTemplates() { 
            var varsToKeep = new List<TemplateVariable>();
            var varsToRemove = new List<TemplateVariable>();

            foreach (string fileName in clbTemplates.Items) {

                var vars = getTemplateVariables(fileName);

                bool isChecked = clbTemplates.CheckedItems.IndexOf(fileName) != -1;

                if (isChecked)
                {
                    varsToKeep.AddRange(vars);
                    varsToKeep = varsToKeep.Distinct().ToList();
                }
                else {
                    varsToRemove.AddRange(vars);
                    varsToRemove = varsToRemove.Distinct().ToList();
                }
            }

            foreach (TemplateVariable varToRemove in varsToRemove) {
                if (!varsToKeep.Contains(varToRemove)) {
                    templateVariableNames.Remove(varToRemove);
                }
            }

            foreach (TemplateVariable varToKeep in varsToKeep) {
                if (!templateVariableNames.Contains(varToKeep)) {
                    templateVariableNames.Add(varToKeep);
                }
            }
        }

        private void scanTemplates() {
            templateVariableNames = new BindingList<TemplateVariable>();

            var templateFilePaths = new BindingList<String>();
            var dirInfo = new DirectoryInfo(templateDirectoryPath);
            foreach (var file in dirInfo.GetFiles())
            {
                if (file.Extension.ToLower().Equals(".docx"))
                {
                    templateFilePaths.Add(file.Name);
                }
            }

            clbTemplates.DataSource = templateFilePaths;

            dgvVariables.DataSource = templateVariableNames;

            if (templateFilePaths.Count == 0) {
                MessageBox.Show("No templates in templates folder", "Warning", MessageBoxButtons.OK);
                System.Diagnostics.Process.Start(templateDirectoryPath);
            }

            for (int i = 0; i < clbTemplates.Items.Count; i++)
            {
                clbTemplates.SetItemChecked(i, true);
            }
            refreshVariablesForAllTemplates();
        }

        private List<TemplateVariable> getTemplateVariables(string fileName) { 
            var result = new List<TemplateVariable>();

            using (var mainDoc = WordprocessingDocument.Open(templateDirectoryPath + @"\" + fileName, false))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(mainDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                var regexName = new Regex(@"{\$(?<name>.*?)\$}");
                var matches = regexName.Matches(docText);

                foreach (Match match in matches) {
                    var name = match.Groups["name"].Value;
                    var variable = new TemplateVariable(name, "");
                    result.Add(variable);
                }
            }
            return result;
        }

        private void generateDocumentFromTemplate(string fileName) {
            using (var mainDoc = WordprocessingDocument.Open(templateDirectoryPath + @"\" + fileName, false))
            using (var resultDoc = WordprocessingDocument.Create(outputDirectoryPath + @"\" + fileName, WordprocessingDocumentType.Document))
            {

                foreach (var part in mainDoc.Parts)
                    resultDoc.AddPart(part.OpenXmlPart, part.RelationshipId);

                string docText = null;
                using (StreamReader sr = new StreamReader(resultDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                foreach (var variable in templateVariableNames) {
                    docText = docText.Replace(String.Format("{{${0}$}}", variable.Name), variable.Value);
                }

                using (StreamWriter sw = new StreamWriter(resultDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgvVariables.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (dgvVariables.Columns[cell.ColumnIndex].HeaderText != "Name") {
                        if (string.IsNullOrEmpty(cell.FormattedValue.ToString()))
                        {
                            MessageBox.Show("Variable value must not be empty");
                            dgvVariables.Rows[cell.RowIndex].ErrorText = "Variable value must not be empty";
                            dgvVariables.CurrentCell = cell;
                            dgvVariables.BeginEdit(true);
                            return;
                        }
                    }
                }
            }

            foreach (string fileName in clbTemplates.Items)
            {
                bool isChecked = clbTemplates.CheckedItems.IndexOf(fileName) != -1;
                if (isChecked)
                {
                    generateDocumentFromTemplate(fileName);
                }
            }

            System.Diagnostics.Process.Start(outputDirectoryPath);
        }
    }

    public class TemplateVariable : IEquatable<TemplateVariable>, IComparable<TemplateVariable>
    {
        public String Name { get; set; }
        public String Value { get; set; }

        public TemplateVariable(string nameArg, string valueArg)
        {
            Name = nameArg;
            Value = valueArg;
        }

        public bool Equals(TemplateVariable other)
        {
            return this.Name.Equals(other.Name);
        }

        public int CompareTo(TemplateVariable other)
        {
           return this.Name.CompareTo(other.Name);
        }
    }
}
