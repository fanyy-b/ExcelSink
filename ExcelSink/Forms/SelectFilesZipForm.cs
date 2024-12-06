using System.Linq;

namespace ExcelSink.Forms
{
    public partial class SelectFilesZipForm : Form
    {
        public SelectFilesZipForm()
        {
            InitializeComponent();
        }

        private IEnumerable<string> _files;

        public void SetFiles(IEnumerable<string> files)
        {
            _files = files;
            UpdateData();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            UpdateData();
        }

        private void UpdateData()
        {
            // Získání textu z TextBoxu
            string filter = textBox1.Text.ToLower();

            // Filtrování položek
            var filteredItems = _files
                .Where(f => f.ToLower().Contains(filter))
                .OrderDescending()
                .ToArray();

            // Aktualizace ListBoxu
            listBox1.Items.Clear();
            listBox1.Items.AddRange(filteredItems);
        }

        public string[] SelectedFiles 
            =>this.listBox1.SelectedItems.Cast<string>().ToArray();
    }
}
