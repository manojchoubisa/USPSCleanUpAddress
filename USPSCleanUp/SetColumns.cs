using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace USPSCleanUp
{
  public partial class SetColumns : Form
    {
        public static int sheetCount;
        public static List<TotalColumnList> FullColumnList;
        public static string name;
        public SetColumns()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {           

            try
            {
               
                
                
                    TotalColumnList obj = new TotalColumnList();
                   obj.SheetName = lblSheetName.Text;
                  if (cbxAddress_Line1.SelectedIndex != 0)
                    {
                        obj.Addressline1col = cbxAddress_Line1.SelectedValue.ToString();
                    }
                    else
                    {
                        obj.Addressline1col = "";
                    }
                    if (cbxAddressLine_2.SelectedIndex != 0)
                    {
                        obj.AddressLine2col = cbxAddressLine_2.SelectedValue.ToString();
                    }
                    else
                    {
                        obj.AddressLine2col = "";

                    }

                    if (cbxzip.SelectedIndex != 0)
                    {
                        obj.Zipcol = cbxzip.SelectedValue.ToString();
                    }
                    else
                    {
                        obj.Zipcol = "";
                    }
                    if (cbxState.SelectedIndex != 0)
                    {
                        obj.Statecol = cbxState.SelectedValue.ToString();
                    }
                    else
                    {
                        obj.Statecol = "";
                    }
                    if (cbxCity.SelectedIndex != 0)
                    {
                        obj.Citycol = cbxCity.SelectedValue.ToString();
                    }
                    else
                    {
                        obj.Citycol = "";
                    }
                    if (cbxHouseNo.SelectedIndex != 0)
                    {
                        obj.Housenocol = cbxHouseNo.SelectedValue.ToString();
                    }
                    else
                    {
                        obj.Housenocol = "";
                    }
                 
                    if (obj.Addressline1col == "")
                    {
                        lblAddress1.Visible = true;
                        lblAddress1.Text = "Address Required";
                        lblAddress1.ForeColor = Color.Red;
                    }
                    else if (obj.Zipcol == "")
                    {
                        lblZipCode.Visible = true;
                        lblZipCode.Text = "Zip Code Required";
                        lblZipCode.ForeColor = Color.Red;
                    }

                    
                    else
                    {
                 
                    FullColumnList.Add(obj);
                    var fileupload = new UploadFile();
                        fileupload.Show();
                        this.Hide();
                    }
                          
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void SetColumns_Load(object sender, EventArgs e)
        {
            try
            {
                FullColumnList = new List<TotalColumnList>();

                sheetCount = 1;
                var SheetNamesList = USPSCleanUp.UploadFile.dtsheetName;
              
                var first = SheetNamesList[0].ToString();
                lblSheetName.Text = first;
                name = first;
                foreach (DataColumn dc in USPSCleanUp.UploadFile.dt.Tables[first].Columns)
                {
                    string collist = dc.ToString();
                    USPSCleanUp.UploadFile.ColumnList.Add(collist);
                }
                //var filename = USPSCleanUp.UploadFile.onlyFileName;
                var ColumnList = USPSCleanUp.UploadFile.ColumnList;
                ColumnList.Insert(0, "--Select--");
                cbxAddress_Line1.DataSource = new BindingSource(ColumnList, null);
                cbxAddressLine_2.DataSource = new BindingSource(ColumnList, null);
                cbxzip.DataSource = new BindingSource(ColumnList, null);
                cbxState.DataSource = new BindingSource(ColumnList, null);
                cbxCity.DataSource = new BindingSource(ColumnList, null);
                cbxHouseNo.DataSource = new BindingSource(ColumnList, null);
                if (sheetCount== SheetNamesList.Count)
                {

                    btnNextSheet.Visible = false;
                    btnSkip.Visible = false;
                }
                else
                {
                    btnRun.Visible = false;

                }

            }
            catch (Exception)
            {
                throw;
            }
        }

        private void cbSheetName_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

            ComboBox comboBox = (ComboBox)sender;

            // Save the selected employee's name, because we will remove
            // the employee's name from the list.
            string Sheetname = (string)comboBox.SelectedItem;
            USPSCleanUp.UploadFile.ColumnList.Clear();
          
    
          
            foreach (DataColumn dc in USPSCleanUp.UploadFile.dt.Tables[Sheetname].Columns)
            {
                string collist = dc.ToString();
                USPSCleanUp.UploadFile.ColumnList.Add(collist);
            }
            //var filename = USPSCleanUp.UploadFile.onlyFileName;
            var ColumnList = USPSCleanUp.UploadFile.ColumnList;
            ColumnList.Insert(0, "--Select--");
            cbxAddress_Line1.DataSource = new BindingSource(ColumnList, null);
            cbxAddressLine_2.DataSource = new BindingSource(ColumnList, null);
            cbxzip.DataSource = new BindingSource(ColumnList, null);
            cbxState.DataSource = new BindingSource(ColumnList, null);
            cbxCity.DataSource = new BindingSource(ColumnList, null);
            cbxHouseNo.DataSource = new BindingSource(ColumnList, null);

            }
            catch (Exception)
            {

                throw;
            }

        }

        private void btnNextSheet_Click(object sender, EventArgs e)
        {
            try
            {
                var SheetNamesList = USPSCleanUp.UploadFile.dtsheetName;
                if (SheetNamesList.Count <= sheetCount + 1)
                {
                   
                    TotalColumnList obj = new TotalColumnList();
                    obj.SheetName = lblSheetName.Text;
                    var first = SheetNamesList[sheetCount].ToString();
                    lblSheetName.Text = first;
                   
                    if (cbxAddress_Line1.SelectedIndex != 0)
                    {
                        obj.Addressline1col = cbxAddress_Line1.SelectedValue.ToString();
                    }
                    else
                    {
                        obj.Addressline1col = "";
                    }
                    if (cbxAddressLine_2.SelectedIndex != 0)
                    {
                        obj.AddressLine2col = cbxAddressLine_2.SelectedValue.ToString();
                    }
                    else
                    {
                        obj.AddressLine2col = "";

                    }

                    if (cbxzip.SelectedIndex != 0)
                    {
                        obj.Zipcol = cbxzip.SelectedValue.ToString();
                    }
                    else
                    {
                        obj.Zipcol = "";
                    }
                    if (cbxState.SelectedIndex != 0)
                    {
                        obj.Statecol = cbxState.SelectedValue.ToString();
                    }
                    else
                    {
                        obj.Statecol = "";
                    }
                    if (cbxCity.SelectedIndex != 0)
                    {
                        obj.Citycol = cbxCity.SelectedValue.ToString();
                    }
                    else
                    {
                        obj.Citycol = "";
                    }
                    if (cbxHouseNo.SelectedIndex != 0)
                    {
                        obj.Housenocol = cbxHouseNo.SelectedValue.ToString();
                    }
                    else
                    {
                        obj.Housenocol = "";
                    }
               

             
                    if (obj.Addressline1col == "")
                    {
                        lblAddress1.Visible = true;
                        lblAddress1.Text = "Address Required";
                        lblAddress1.ForeColor = Color.Red;
                    }
                    else if (obj.Zipcol == "")
                    {
                        lblZipCode.Visible = true;
                        lblZipCode.Text = "Zip Code Required";
                        lblZipCode.ForeColor = Color.Red;
                    }
                    USPSCleanUp.UploadFile.ColumnList.Clear();
                    foreach (DataColumn dc in USPSCleanUp.UploadFile.dt.Tables[first].Columns)
                  {
                        string collist = dc.ToString();
                        USPSCleanUp.UploadFile.ColumnList.Add(collist);
                    }
                    //var filename = USPSCleanUp.UploadFile.onlyFileName;
                    var ColumnList = USPSCleanUp.UploadFile.ColumnList;
                    ColumnList.Insert(0, "--Select--");
                    cbxAddress_Line1.DataSource = new BindingSource(ColumnList, null);
                    cbxAddressLine_2.DataSource = new BindingSource(ColumnList, null);
                    cbxzip.DataSource = new BindingSource(ColumnList, null);
                    cbxState.DataSource = new BindingSource(ColumnList, null);
                    cbxCity.DataSource = new BindingSource(ColumnList, null);
                    cbxHouseNo.DataSource = new BindingSource(ColumnList, null);
                    sheetCount += 1;
                    if (sheetCount == SheetNamesList.Count)
                    {
                        if (sheetCount == 2)
                        {
                            btnSkip.Visible = true;
                           
                        }
                        btnRun.Visible = true;
                        btnNextSheet.Visible = false;
                       
                    }
                    else
                    {

                        if (sheetCount == 2)
                        {
                            btnSkip.Visible = true;
                        }
                        btnRun.Visible = false;
                        btnNextSheet.Visible = true;
                      
                    }
                 
                    FullColumnList.Add(obj);
                }
         
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            USPSCleanUp.UploadFile.upload = false;
            USPSCleanUp.UploadFile.dtsheetName.Clear();
            USPSCleanUp.SetColumns.sheetCount = 0;
            USPSCleanUp.SetColumns.FullColumnList.Clear();
            USPSCleanUp.UploadFile.dt.Clear();
            
            var fileupload = new UploadFile();
            fileupload.Show();
            this.Hide();
        }

        private void btnSkip_Click(object sender, EventArgs e)
        {
            try
            {
                
            
                var SheetNamesList = USPSCleanUp.UploadFile.dtsheetName;
                sheetCount = SheetNamesList.Count;

                btnNextSheet.Visible = false;
                btnSkip.Visible = false;
                btnRun.Visible = true;
                foreach (var col in SheetNamesList.ToList())
                {
                    if (col!= name)
                    {
                        SheetNamesList.Remove(col);
                        if (USPSCleanUp.UploadFile.dt.Tables.Contains(col) )
                            USPSCleanUp.UploadFile.dt.Tables.Remove(col);
                
                    }

                  
                }



                var fileupload = new UploadFile();
                    fileupload.Show();
                    this.Hide();
               
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
