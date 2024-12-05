using ExcelSink.Forms;
using System.Data;
using System.Globalization;
using System.IO.Compression;

namespace ExcelSink.Model
{
    public class MeasurementsDataTable : DataTable
    {
        public MeasurementsDataTable()
        {
            this.Columns.Add("Date", typeof(DateOnly));
            this.Columns.Add("SensorId", typeof(string));
            this.Columns.Add("Power", typeof(decimal));
            this.Columns.Add("Temperature", typeof(decimal));
            this.Columns.Add("Time ", typeof(MyTimeOnly));
        }

        public int LoadCsvFile(string fileName)
        {
            try
            {
                var s = File.OpenRead(fileName);
                return LoadCsvStream(s, fileName);
            }
            catch
            {
                return -1;
            }
        }

        public int LoadCsvStream(Stream stream, string fileName)
        {
            try
            {
                var sr = new StreamReader(stream);
                var lines = sr.ReadToEnd().Split(new char[] { '\n' });

                var tempDataTable = new List<MeasureRow>();

                foreach (var line in lines)
                {
                    var parts = line.Split(';');
                    if (parts.Length != 6)
                    {
                        MessageBox.Show($"Soubor {fileName} nemá správný formát!", "Varování", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return -2;
                    }

                    var dateParts = parts[2].Split('.');
                    if (dateParts.Length != 3)
                    {
                        MessageBox.Show($"Soubor {fileName} obsahuje nesprávně formátované datum: '{parts[2]}'.", "Varování", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return -2;
                    }
                    var date = new DateOnly(int.Parse(dateParts[0]), int.Parse(dateParts[1]), int.Parse(dateParts[2]));

                    var timeParts = parts[3].Split(':');
                    if (timeParts.Length != 3)
                    {
                        MessageBox.Show($"Soubor {fileName} obsahuje nesprávně formátovaný čas: '{parts[2]}'.", "Varování", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return -2;
                    }
                    var time = new MyTimeOnly(int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse(timeParts[2]));


                    tempDataTable.Add(
                        new MeasureRow(
                            parts[0],
                            date,
                            time,
                            decimal.Parse(parts[4], CultureInfo.InvariantCulture.NumberFormat),
                            decimal.Parse(parts[5], CultureInfo.InvariantCulture.NumberFormat)));
                }

                foreach (var dr in tempDataTable)
                    Rows.Add(dr.Date, dr.SensorId, dr.Power, dr.Temperature, dr.Time);

                return Rows.Count;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }
        }


        public int LoadZipFile(string fileName)
        {
            var stream = File.OpenRead(fileName);
            ZipArchive za = new ZipArchive(stream);

            var selectZipForm = new SelectFilesZipForm();
            selectZipForm.SetFiles(za.Entries
                .Where(z => Path.GetExtension(z.Name) == ".txt")
                .Select(z => z.Name));

            if (selectZipForm.ShowDialog() == DialogResult.OK)
            {
                foreach (var entry in za.Entries.Where(e => selectZipForm.SelectedFiles.Contains(e.Name)))
                {
                    var s = entry.Open();
                    var ms = new MemoryStream();
                    StreamReader sr = new StreamReader(s);

                    s.CopyTo(ms);
                    ms.Seek(0, SeekOrigin.Begin);

                    LoadCsvStream(ms, entry.Name);
                }
            }

            return -1;
        }
    }
}

