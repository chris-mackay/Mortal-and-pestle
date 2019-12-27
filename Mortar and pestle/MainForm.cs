using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Mortar_and_pestle
{
    public partial class MainForm : Form
    {
        OleDbConnection con = new OleDbConnection();
        string provider = string.Empty;
        string source = string.Empty;
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        OleDbDataAdapter da;
        string sql = string.Empty;

        private string file = "Ingredients.accdb";
        public string select = "SELECT ID, Ingredient, Weight, Value, PrimaryEffect, SecondaryEffect, TertiaryEffect, QuaternaryEffect FROM tblIngredients";

        public MainForm()
        {
            InitializeComponent();
        }

        private void btnFilterSelected_Click(object sender, EventArgs e)
        {
            filterDataExact(dgvIngredients.CurrentCell.Value.ToString());
        }

        private void btnViewAll_Click(object sender, EventArgs e)
        {
            DrawingControl.SetDoubleBuffered(dgvIngredients);
            DrawingControl.SuspendDrawing(dgvIngredients);

            txtFilter.Text = string.Empty;
            DBConnect();
            SQLSelect(select);
            Fill_Format_DataGrid();
            DBClose();

            dgvIngredients.ClearSelection();
            DrawingControl.ResumeDrawing(dgvIngredients);
        }

        private void txtFilter_TextChanged(object sender, EventArgs e)
        {
            filterDataThatContainsItemOnTextChanged();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            DrawingControl.SetDoubleBuffered(dgvIngredients);
            DrawingControl.SuspendDrawing(dgvIngredients);

            DBConnect();
            SQLSelect(select);
            Fill_Format_DataGrid();
            DBClose();

            dgvIngredients.ClearSelection();
            DrawingControl.ResumeDrawing(dgvIngredients);
        }

        public void filterDataExact(string _item)
        {
            try
            {
                DrawingControl.SetDoubleBuffered(dgvIngredients);
                DrawingControl.SuspendDrawing(dgvIngredients);

                txtFilter.Text = _item;

                OleDbConnection con = new OleDbConnection();
                OleDbDataAdapter da = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                BindingSource bs = new BindingSource();
                DataView dsView = new DataView();
                
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + file;
                con.Open();
                ds.Tables.Add(dt);
                da = new OleDbDataAdapter("SELECT * FROM tblIngredients", con);
                da.Fill(dt);
                dsView = ds.Tables[0].DefaultView;
                bs.DataSource = dsView;
                _item = _item.Replace("'", "").Replace("_", "").Replace("%", "");
                string strSearch = txtFilter.Text.Replace("'", "").Replace("_", "").Replace("%", "");

                bs.Filter = " Ingredient LIKE'%" + strSearch + "%' " +
                            "or PrimaryEffect LIKE'%" + strSearch + "%' " +
                            "or SecondaryEffect LIKE'%" + strSearch + "%' " +
                            "or TertiaryEffect LIKE'%" + strSearch + "%' " +
                            "or QuaternaryEffect LIKE'%" + strSearch + "%'" +
                            "or Value LIKE'%" + strSearch + "%'" +
                            "or Weight LIKE'%" + strSearch + "%'";

                dgvIngredients.DataSource = bs;
                con.Close();

                dgvIngredients.ClearSelection();
                DrawingControl.ResumeDrawing(dgvIngredients);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Connection could not be made\n\n" + ex.Message + "\n\n" + ex.Source, "Connection Failure", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        public void filterDataThatContainsItemOnTextChanged()
        {
            try
            {
                DrawingControl.SetDoubleBuffered(dgvIngredients);
                DrawingControl.SuspendDrawing(dgvIngredients);
                
                OleDbConnection con = new OleDbConnection();
                OleDbDataAdapter da = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                BindingSource bs = new BindingSource();
                DataView dsView = new DataView();
                
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + file;
                con.Open();
                ds.Tables.Add(dt);
                da = new OleDbDataAdapter("SELECT * FROM tblIngredients", con);
                da.Fill(dt);
                dsView = ds.Tables[0].DefaultView;
                bs.DataSource = dsView;
                string strSearch = txtFilter.Text.Replace("'", "").Replace("_", "").Replace("%", "");

                bs.Filter = " Ingredient LIKE'%" + strSearch + "%' " +
                            "or PrimaryEffect LIKE'%" + strSearch + "%' " +
                            "or SecondaryEffect LIKE'%" + strSearch + "%' " +
                            "or TertiaryEffect LIKE'%" + strSearch + "%' " +
                            "or QuaternaryEffect LIKE'%" + strSearch + "%'" +
                            "or Value LIKE'%" + strSearch + "%'" +
                            "or Weight LIKE'%" + strSearch + "%'";

                dgvIngredients.DataSource = bs;
                con.Close();

                dgvIngredients.ClearSelection();
                DrawingControl.ResumeDrawing(dgvIngredients);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Connection could not be made\n\n" + ex.Message + "\n\n" + ex.Source, "Connection Failure", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        public void DBConnect()
        {
            con.Close();
            provider = "PROVIDER=Microsoft.ACE.OLEDB.12.0;";
            source = "Data Source = " + file;
            con.ConnectionString = provider + source;
            con.Open();
        }

        public void DBClose()
        {
            con.Close();
        }

        public DataSet SQLSelect(string sqlString)
        {
            sql = sqlString;
            da = new OleDbDataAdapter();
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = con;
            cmd.CommandText = sql;
            da.SelectCommand = cmd;

            da.Fill(ds, "tblIngredients");

            return ds;
        }

        public void Fill_Format_DataGrid()
        {
            dt.Clear();
            da.Fill(dt);
            this.dgvIngredients.DataSource = dt.DefaultView;

            {
                var withBlock = this.dgvIngredients;

                withBlock.Columns["ID"].Visible = false;

                withBlock.Columns["Ingredient"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                withBlock.Columns["Ingredient"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                withBlock.Columns["Ingredient"].Width = 150;

                withBlock.Columns["Weight"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                withBlock.Columns["Weight"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                withBlock.Columns["Weight"].Width = 50;

                withBlock.Columns["Value"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                withBlock.Columns["Value"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                withBlock.Columns["Value"].Width = 50;

                withBlock.Columns["PrimaryEffect"].HeaderText = "Primary Effect";
                withBlock.Columns["PrimaryEffect"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                withBlock.Columns["PrimaryEffect"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                withBlock.Columns["PrimaryEffect"].Width = 150;

                withBlock.Columns["SecondaryEffect"].HeaderText = "Secondary Effect";
                withBlock.Columns["SecondaryEffect"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                withBlock.Columns["SecondaryEffect"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                withBlock.Columns["SecondaryEffect"].Width = 150;

                withBlock.Columns["TertiaryEffect"].HeaderText = "Tertiary Effect";
                withBlock.Columns["TertiaryEffect"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                withBlock.Columns["TertiaryEffect"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                withBlock.Columns["TertiaryEffect"].Width = 150;

                withBlock.Columns["QuaternaryEffect"].HeaderText = "Quaternary Effect";
                withBlock.Columns["QuaternaryEffect"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                withBlock.Columns["QuaternaryEffect"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            }
        }

        private void dgvIngredients_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (txtFilter.Text != string.Empty && e.Value.ToString().ToUpper().Contains(txtFilter.Text.ToUpper()))
                dgvIngredients.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Aquamarine;
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            dgvIngredients.ClearSelection();
        }
    }

    public static class DrawingControl
    {
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr _hWnd, Int32 _wMsg, bool _wParam, Int32 _lParam);

        private const int WM_SETREDRAW = 11;

        public static void SetDoubleBuffered(Control _ctrl)
        {
            if (!SystemInformation.TerminalServerSession)
            {
                typeof(Control).InvokeMember("DoubleBuffered", (System.Reflection.BindingFlags.SetProperty
                                | (System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)), null, _ctrl, new object[] {
                            true});
            }
        }

        public static void SetDoubleBuffered_ListControls(List<Control> _ctrlList)
        {
            if (!SystemInformation.TerminalServerSession)
            {
                foreach (Control ctrl in _ctrlList)
                {
                    typeof(Control).InvokeMember("DoubleBuffered", (System.Reflection.BindingFlags.SetProperty
                                    | (System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)), null, ctrl, new object[] {
                                true});
                }
            }
        }

        public static void SuspendDrawing(Control _ctrl)
        {
            SendMessage(_ctrl.Handle, WM_SETREDRAW, false, 0);
        }

        public static void SuspendDrawing_ListControls(List<Control> _ctrlList)
        {
            foreach (Control ctrl in _ctrlList)
            {
                SendMessage(ctrl.Handle, WM_SETREDRAW, false, 0);
            }
        }

        public static void ResumeDrawing(Control _ctrl)
        {
            SendMessage(_ctrl.Handle, WM_SETREDRAW, true, 0);
            _ctrl.Refresh();
        }

        public static void ResumeDrawing_ListControls(List<Control> _ctrlList)
        {
            foreach (Control ctrl in _ctrlList)
            {
                SendMessage(ctrl.Handle, WM_SETREDRAW, true, 0);
                ctrl.Refresh();
            }
        }
    }
}

