using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Threading;
using System.Diagnostics;

namespace DOCXTemplate
{
    public partial class Main : Form
    {

        BindingList<TemplateVariable> templateVariableNames = new BindingList<TemplateVariable>();
        BindingList<String> templateFilePaths = new BindingList<String>();
        String templateDirectoryPath = AppDomain.CurrentDomain.BaseDirectory + @"\templates";
        String templateSubDirectoryPath = AppDomain.CurrentDomain.BaseDirectory + @"\templates\sub";
        String outputDirectoryPath = AppDomain.CurrentDomain.BaseDirectory + @"\output";
        BindingList<String> subTemplateFilePaths = new BindingList<String>();
        Dictionary<string, string> subTemplatesUpdatedDocuments = new Dictionary<string, string>();

        LoadingDialog dialog = new LoadingDialog();

        public Main()
        {
            InitializeComponent();

            if (!Directory.Exists(templateDirectoryPath))
            {
                Directory.CreateDirectory(templateDirectoryPath);
            }

            if (!Directory.Exists(templateSubDirectoryPath))
            {
                Directory.CreateDirectory(templateSubDirectoryPath);
            }

            if (!Directory.Exists(outputDirectoryPath))
            {
                Directory.CreateDirectory(outputDirectoryPath);
            }

            this.Shown += Main_Shown;

            clbTemplates.DataSource = templateFilePaths;
        }

        private void Main_Shown(object sender, EventArgs e)
        {
            _ = scanTemplatesTask();
        }

        private void btnScanTemplates_Click(object sender, EventArgs e)
        {
            _ = scanTemplatesTask();
        }

        private async Task scanTemplatesTask()
        {
            await Task.Run(() => {
                this.BeginInvoke(new MethodInvoker(delegate {
                    dialog.Show(this);
                }));

                scanTemplates();

                Thread.Sleep(500);

                this.BeginInvoke(new MethodInvoker(delegate {
                    dialog.Hide();
                }));
            });
        }
        private void scanTemplates()
        {
            templateFilePaths.Clear();
            subTemplateFilePaths.Clear();
            var dirInfo = new DirectoryInfo(templateDirectoryPath);
            foreach (var file in dirInfo.GetFiles())
            {
                if (file.Extension.ToLower().Equals(".docx"))
                {
                    templateFilePaths.Add(file.Name);
                }
            }

            dirInfo = new DirectoryInfo(templateSubDirectoryPath);
            foreach (var file in dirInfo.GetFiles())
            {
                if (file.Extension.ToLower().Equals(".docx"))
                {
                    templateFilePaths.Add(@"sub\" + file.Name);
                    subTemplateFilePaths.Add(file.Name);
                }
            }

            if (clbTemplates.InvokeRequired)
            {
                clbTemplates.Invoke(new MethodInvoker(delegate
                {
                    templateFilePaths.ResetBindings();
                    for (int i = 0; i < clbTemplates.Items.Count; i++)
                    {
                        clbTemplates.SetItemChecked(i, true);
                    }
                }));
            }
            else {
                templateFilePaths.ResetBindings();
                for (int i = 0; i < clbTemplates.Items.Count; i++)
                {
                    clbTemplates.SetItemChecked(i, true);
                }
            }

            if (this.InvokeRequired)
            {
                this.BeginInvoke(new MethodInvoker(delegate
                {
                    if (templateFilePaths.Count == 0)
                    {
                        MessageBox.Show("No templates in templates folder", "Warning", MessageBoxButtons.OK);
                        Process.Start(templateDirectoryPath);
                    }
                }));
            }
            else { 
                                if (templateFilePaths.Count == 0)
                    {
                        MessageBox.Show("No templates in templates folder", "Warning", MessageBoxButtons.OK);
                        Process.Start(templateDirectoryPath);
                    }
            }
        }
        private void clbTemplates_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke(new MethodInvoker(delegate
            {
                var displayLoadingDialog = !dialog.Visible;
                _ = refreshVariablesForAllTemplatesTask(displayLoadingDialog);
            }));
        }
        private void btnToggleSelect_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < clbTemplates.Items.Count; i++) {
                var newChecked = !(clbTemplates.GetItemCheckState(i) == CheckState.Checked);
                clbTemplates.SetItemChecked(i, newChecked);
            }
        }

        private async Task refreshVariablesForAllTemplatesTask(bool showHideLoadingDialog)
        {
            await Task.Run(() => {
                if (showHideLoadingDialog)
                {
                    this.BeginInvoke(new MethodInvoker(delegate
                    {
                        dialog.Show(this);
                    }));
                }

                refreshVariablesForTemplate();

                Thread.Sleep(500);

                if (showHideLoadingDialog)
                {
                    this.BeginInvoke(new MethodInvoker(delegate
                    {
                        dialog.Hide();
                    }));
                }
            });
        }

        private void refreshVariablesForTemplate() {

            var updatedVariables = new Dictionary<string, TemplateVariable>();
            foreach (TemplateVariable variable in templateVariableNames)
            {
                updatedVariables.Add(variable.Name, variable);
            }

            foreach (string fileName in templateFilePaths)
            {
                var indexFile = clbTemplates.CheckedItems.IndexOf(fileName);
                bool isChecked = indexFile != -1;

                if (fixTemplate(fileName))
                {
                    String message = String.Format("Template {0} had broken styles in variable. It was fixed. The backup file save to {0}.bak", fileName);
                    if (this.InvokeRequired)
                    {
                        this.BeginInvoke(new MethodInvoker(delegate
                        {
                            MessageBox.Show(message, "Warning", MessageBoxButtons.OK);
                        }));
                    }
                    else
                    {
                        MessageBox.Show(message, "Warning", MessageBoxButtons.OK);
                    }
                }

                var variablesInTemplate = getTemplateVariables(fileName);

                if (isChecked)
                {
                    foreach (TemplateVariable variable in variablesInTemplate) {
                        if (updatedVariables.ContainsKey(variable.Name))
                        {
                            updatedVariables[variable.Name].InTemplates.Add(fileName);
                        }
                        else {
                            variable.InTemplates.Add(fileName);
                            updatedVariables.Add(variable.Name, variable);
                        }
                    }
                }
                else
                {
                    foreach (TemplateVariable variable in variablesInTemplate)
                    {
                        if (updatedVariables.ContainsKey(variable.Name))
                        {
                            updatedVariables[variable.Name].InTemplates.Remove(fileName);
                            if (updatedVariables[variable.Name].InTemplates.Count == 0) {
                                updatedVariables.Remove(variable.Name);
                            }
                        }
                    }
                }
            }
            if (dgvVariables.InvokeRequired)
            {
                dgvVariables.BeginInvoke(new MethodInvoker(delegate
                {
                    updateVariableGridView(updatedVariables);
                }));
            }
            else
            {
                updateVariableGridView(updatedVariables);
            }
        }

        private void updateVariableGridView(Dictionary<string, TemplateVariable> updatedVariables)
        {
            templateVariableNames = new BindingList<TemplateVariable>(updatedVariables.Values.OrderBy(v => v.Name).ToList());
            dgvVariables.DataSource = templateVariableNames;
        }

        private bool fixTemplate(string fileName) {
            bool result = false;
            var templatePath = templateDirectoryPath + @"\" + fileName;
            List<string> tagsToReplace = new List<string>();
            using (var mainDoc = WordprocessingDocument.Open(templatePath, false))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(mainDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                var regexVariable = new Regex(@"{(?<capture>(?:\<.*?\>)?\$(?:\<.*?\>)?.*?(?:\<.*?\>)?\$(?:\<.*?\>)?)}");
                var matches = regexVariable.Matches(docText);

                foreach (Match match in matches)
                {
                    String capture = match.ToString();
                    var variableTag = Regex.Replace(capture, @"\<.*?\>", "");
                    if (!capture.Equals(variableTag))
                    {
                        tagsToReplace.Add(capture);
                    }
                }
            }

            result = tagsToReplace.Count > 0;

            if (result)
            {
                File.Copy(templatePath, templatePath + ".bak", true);
                using (var mainDoc = WordprocessingDocument.Open(templatePath, true))
                {
                    string docText = null;
                    using (StreamReader sr = new StreamReader(mainDoc.MainDocumentPart.GetStream()))
                    {
                        docText = sr.ReadToEnd();
                    }

                    foreach (String tagToReplace in tagsToReplace)
                    {
                        var variableTag = Regex.Replace(tagToReplace, @"\<.*?\>", "");
                        docText = docText.Replace(tagToReplace, variableTag);
                    }

                    using (StreamWriter sw = new StreamWriter(mainDoc.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }
                }
            }
            return result;
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
                    var variable = new TemplateVariable(name, "", TemplateVariableType.Text);
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
                    if (variable.Value.Trim().Length != 0)
                    {
                        docText = docText.Replace(String.Format("{{${0}$}}", variable.Name), variable.Value);
                    }
                }

                using (StreamWriter sw = new StreamWriter(resultDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        private bool hasSubTemplateVariables(string fileName)
        {
            bool result = templateVariableNames.Any(x => x.InTemplates.Any(t => t.StartsWith(@"sub/")));
            return result;
        }

        private async Task generateDocumentFromTemplateCheckedTask()
        {
            await Task.Run(() => {

                this.BeginInvoke(new MethodInvoker(delegate
                {
                    dialog.Show(this);
                }));

                subTemplatesUpdatedDocuments.Clear();

                List<string> templatesChecked = new List<string>();
                List<string> subTemplatesChecked = new List<string>();
                foreach (string fileName in clbTemplates.CheckedItems)
                {
                    if (fileName.StartsWith(@"sub/"))
                    {
                        subTemplatesChecked.Add(fileName);
                    }
                    else {
                        templatesChecked.Add(fileName);
                    }
                }

                foreach (string fileName in subTemplatesChecked)
                {
                    //generateDocumentFromTemplate(fileName);
                }

                foreach (string fileName in templatesChecked)
                {
                    generateDocumentFromTemplate(fileName);
                }

                Thread.Sleep(500);

                this.BeginInvoke(new MethodInvoker(delegate
                {
                    dialog.Hide();
                    Process.Start(outputDirectoryPath);
                }));
            });
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            _ = generateDocumentFromTemplateCheckedTask();
        }

        private void btnTemplates_Click(object sender, EventArgs e)
        {
           Process.Start(templateDirectoryPath);
        }

        private void btnOutput_Click(object sender, EventArgs e)
        {
            Process.Start(outputDirectoryPath);
        }

        private void dgvVariables_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            ComboBox combo = e.Control as ComboBox;
            if (combo != null)
            {
                // Remove an existing event-handler, if present, to avoid 
                // adding multiple handlers when the editing control is reused.
                combo.SelectedIndexChanged -= dgvVariables_ComboBox_SelectedIndexChanged;

                // Add the event handler. 
                combo.SelectedIndexChanged += dgvVariables_ComboBox_SelectedIndexChanged;
            }
        }
        private void dgvVariables_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox combo = (ComboBox)sender;
            int rowIndex = dgvVariables.CurrentCell.RowIndex;

            var selectedItem = combo.SelectedItem;
            if (null != selectedItem)
            {
                if (selectedItem.GetType() == typeof(TemplateVariableType))
                {
                    TemplateVariableType varType = (TemplateVariableType)selectedItem;
                    switch (varType)
                    {
                        case TemplateVariableType.Text:

                            DataGridViewTextBoxCell TextBoxCell = new DataGridViewTextBoxCell();
                            dgvVariables[2, rowIndex] = TextBoxCell;
                            dgvVariables[2, rowIndex].Value = "";

                            break;
                        case TemplateVariableType.SubTemplate:

                            DataGridViewComboBoxCell ComboBoxCell = new DataGridViewComboBoxCell();
                            ComboBoxCell.DataSource = subTemplateFilePaths;
                            dgvVariables[2, rowIndex] = ComboBoxCell;

                            break;
                        default:
                            break;
                    }
                }
            }
        }
    }

    public enum TemplateVariableType
    {
        Text,
        SubTemplate
    }

    public class TemplateVariable : IEquatable<TemplateVariable>, IComparable<TemplateVariable>
    {
        public String Name { get; set; }
        public String Value { get; set; }
        public TemplateVariableType Type { get; set; }
        public HashSet<string> InTemplates { get; set; }

        public TemplateVariable(string nameArg, string valueArg, TemplateVariableType typeArg)
        {
            Name = nameArg;
            Value = valueArg;
            Type = typeArg;
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
