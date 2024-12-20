using Inventor;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace InventorAddButtonAPI
{
    [Guid("C960A3CB-2DFB-4B53-A52F-C5508F26D0DE")]
    [ComVisible(true)]
    public class AddIn : ApplicationAddInServer
    {
        private static string logFile = @"E:\PreranaWorkplace\Autodesk Inventor\Inventor API\InventorAddButtonAPI\InventorAddInLog.txt";

        private Inventor.Application _inventorApp;
        private ButtonDefinition myButton;
        private RibbonPanel customPanel;

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            _inventorApp = addInSiteObject.Application;
            Ribbon partRibbon = _inventorApp.UserInterfaceManager.Ribbons["Part"];
            RibbonTab assembleTab = partRibbon.RibbonTabs["id_TabTools"];
            customPanel = assembleTab.RibbonPanels.Add("Views", "CustomPanel", "");

            myButton = _inventorApp.CommandManager.ControlDefinitions.AddButtonDefinition("Export Views ", "Views ", CommandTypesEnum.kNonShapeEditCmdType);
            customPanel.CommandControls.AddButton(myButton);
            myButton.OnExecute += MyButton_OnExecute;
        }

        private void MyButton_OnExecute(NameValueMap context)
        {
            try
            {
                if (!(_inventorApp.ActiveDocument is PartDocument partDocument))
                {
                    Log("Not a part document.");
                    return;
                }

                // Show checkbox dialog
                var dialog = new ViewSelectionDialog();
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var selectedViews = dialog.SelectedViews;
                    CreateDrawing(partDocument, selectedViews);
                }
            }
            catch (Exception ex)
            {
                Log($"Error: {ex.Message}");
            }
        }

        private void CreateDrawing(PartDocument partDocument, bool[] selectedViews)
        {
            DrawingDocument drawingDoc = (DrawingDocument)_inventorApp.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject);
            Sheet sheet = drawingDoc.Sheets[1];
            double scale = 0.5;

            if (selectedViews[0]) AddView(sheet, partDocument, _inventorApp.TransientGeometry.CreatePoint2d(5.0, 4.0), ViewOrientationTypeEnum.kTopViewOrientation, scale);
            if (selectedViews[1]) AddView(sheet, partDocument, _inventorApp.TransientGeometry.CreatePoint2d(15.0, 4.0), ViewOrientationTypeEnum.kBottomViewOrientation, scale);
            if (selectedViews[2]) AddView(sheet, partDocument, _inventorApp.TransientGeometry.CreatePoint2d(25.0, 4.0), ViewOrientationTypeEnum.kFrontViewOrientation, scale);
            if (selectedViews[3]) AddView(sheet, partDocument, _inventorApp.TransientGeometry.CreatePoint2d(5.0, 14.0), ViewOrientationTypeEnum.kBackViewOrientation, scale);
            if (selectedViews[4]) AddView(sheet, partDocument, _inventorApp.TransientGeometry.CreatePoint2d(15.0, 14.0), ViewOrientationTypeEnum.kLeftViewOrientation, scale);
            if (selectedViews[5]) AddView(sheet, partDocument, _inventorApp.TransientGeometry.CreatePoint2d(25.0, 14.0), ViewOrientationTypeEnum.kRightViewOrientation, scale);

            string outputFile = @"E:\PreranaWorkplace\Autodesk Inventor\Inventor API\InventorAddButtonAPI\Inventor_Exports\output.dwg";
            drawingDoc.SaveAs(outputFile, true);
        }

        private void AddView(Sheet sheet, PartDocument partDocument, Point2d point, ViewOrientationTypeEnum viewOrientation, double scale)
        {
            sheet.DrawingViews.AddBaseView(
                (Inventor._Document)partDocument,
                point,
                scale,
                viewOrientation,
                DrawingViewStyleEnum.kHiddenLineDrawingViewStyle
            );
        }

        private void Log(string message)
        {
            try
            {
                if (System.IO.File.Exists(logFile))
                {
                    using (FileStream fs = System.IO.File.Create(logFile))
                    {
                        fs.Close();
                    }
                }
                using (StreamWriter sw = new StreamWriter(logFile, true))
                {
                    sw.WriteLine($"{message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{ex.Message}");
            }
        }

        public void Deactivate()
        {
            if (_inventorApp != null) RemoveRibbonButton();
            _inventorApp = null;
        }

        private void RemoveRibbonButton()
        {
            try
            {
                UserInterfaceManager uiManager = (UserInterfaceManager)_inventorApp.UserInterfaceManager;
                Ribbon partRibbon = uiManager.Ribbons["Part"];
                RibbonTab assembleTab = partRibbon.RibbonTabs["id_TabTools"];

                foreach (RibbonPanel panel in assembleTab.RibbonPanels)
                {
                    if (panel.DisplayName == "Views")
                    {
                        panel.Delete();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Error during cleanup: " + ex.Message);
            }
        }

        public void ExecuteCommand(int commandID)
        {
            throw new NotImplementedException();
        }

        public object Automation => null;
    }

    // Windows Forms dialog to select views
    public class ViewSelectionDialog : Form
    {
        public bool[] SelectedViews { get; private set; } = new bool[6];
        private CheckBox[] checkBoxes;

        public ViewSelectionDialog()
        {
            Text = "Select Views";
            checkBoxes = new CheckBox[6]
            {
                new CheckBox { Text = "Top View", Location = new System.Drawing.Point(10, 10) },
                new CheckBox { Text = "Bottom View", Location = new System.Drawing.Point(10, 40) },
                new CheckBox { Text = "Front View", Location = new System.Drawing.Point(10, 70) },
                new CheckBox { Text = "Back View", Location = new System.Drawing.Point(10, 100) },
                new CheckBox { Text = "Left View", Location = new System.Drawing.Point(10, 130) },
                new CheckBox { Text = "Right View", Location = new System.Drawing.Point(10, 160) }
            };

            foreach (var checkBox in checkBoxes)
            {
                Controls.Add(checkBox);
            }

            Button okButton = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(10, 190),
                DialogResult = DialogResult.OK
            };
            okButton.Click += (sender, e) =>
            {
                for (int i = 0; i < checkBoxes.Length; i++)
                {
                    SelectedViews[i] = checkBoxes[i].Checked;
                }
            };

            Controls.Add(okButton);
        }
    }
}
