using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;


// shortcut for resources
using res = WordPresent.Properties.Resources;
using System.Data;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace WordPresent
{
    public partial class Ribbon1
    {
        Microsoft.Office.Interop.Word.Application App;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // update ui
            Globals.Ribbons.Ribbon1.btnDataBase.Label = res.btnDataBase;
            Globals.Ribbons.Ribbon1.btnDataBase.ScreenTip = res.btnDataBase_tip;
            Globals.Ribbons.Ribbon1.btnDataBase.SuperTip = res.btnDataBase_tip_content;

            App = Globals.ThisAddIn.Application;
        }

        private void Database_Click(object sender, RibbonControlEventArgs e)
        {
            // select a access database file
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if(dlg.ShowDialog() == DialogResult.OK)
            {
                // store file name
                DataBase.instance.filePath = dlg.FileName;
                DataBase.instance.path = System.IO.Path.GetDirectoryName(dlg.FileName);
                DataBase.instance.fileName = dlg.FileName.Split('\\').Last();
                
                
                // get tables name
                string[] tables = DataBase.instance.GetTableName();
                if (tables == null) return;

                // show the file name
                (sender as RibbonButton).Label = DataBase.instance.fileName;


                if(!tables.Contains("Option$") || !tables.Contains("Format$"))
                    MessageBox.Show(res.msgDefaultTableNoFound);

                // add tables name to combobox 
                Globals.Ribbons.Ribbon1.cmbTables.Items.Clear();
                foreach (string name in tables)
                {
                    if(name != "Option$" && name != "Format$")
                    {
                        RibbonDropDownItem item = this.Factory.CreateRibbonDropDownItem();
                        item.Label = name;
                        Globals.Ribbons.Ribbon1.cmbTables.Items.Add(item);
                    }
                }
            }
        }


        private void Present_Click(object sender, RibbonControlEventArgs e)
        {
            // get data for select table
            DataBase.instance.GetDataTable();
            DataBase.instance.GetOption();
            DataBase.instance.GetFormat();

            string img_dir = DataBase.instance.optionDictionary["img_dir"];
            if (!Path.IsPathRooted(img_dir))
                img_dir = Path.Combine(DataBase.instance.path, img_dir);


            Globals.ThisAddIn.Application.Selection.InsertAfter("\n");

            foreach(Data d in DataBase.instance.dataList)
            {
                string type = d.type;
                if(type == null) continue;

                switch (type.ToLower())
                {                
                    case "txt":
                        Globals.ThisAddIn.Application.Selection.InsertAfter(d.data + '\n');
                        App.Selection.set_Style(DataBase.instance.formatDictionary[d.format].style);
                        Globals.ThisAddIn.Application.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        break;

                    case "img":
                        var shape = Globals.ThisAddIn.Application.Selection.InlineShapes.AddPicture(Path.Combine(img_dir,d.data));

                        // set img size
                        shape.Width =App.InchesToPoints(DataBase.instance.formatDictionary[d.format].width);
                        shape.Height=App.InchesToPoints(DataBase.instance.formatDictionary[d.format].height);

                        break;
                    
                    case "br":
                        App.Selection.InsertAfter("\n");
                        App.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        break;
                    case "end":
                        return;
                }
            }
        }

        private void cmbTables_TextChanged(object sender, RibbonControlEventArgs e)
        {
            // store table name
            DataBase.instance.selectTableName = (sender as RibbonComboBox).Text;
        }
    }
}
