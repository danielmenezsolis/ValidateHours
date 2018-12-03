using System;
using System.AddIn;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using RightNow.AddIns.AddInViews;
using RightNow.AddIns.Common;

namespace ValidateHours
{
    public class WorkspaceAddIn : Panel, IWorkspaceComponent2
    {
        public IGlobalContext gContext { get; set; }

        private IRecordContext _recordContext;


        public WorkspaceAddIn(bool inDesignMode, IRecordContext RecordContext, IGlobalContext global)
        {
            if (!inDesignMode)
            {
                _recordContext = RecordContext;
                gContext = global;
            }
        }

        #region IAddInControl Members


        public Control GetControl()
        {
            return this;
        }

        #endregion

        #region IWorkspaceComponent2 Members

        /// <summary>
        /// Sets the ReadOnly property of this control.
        /// </summary>
        public bool ReadOnly { get; set; }
        public IGenericObject Itinerary { get; set; }
        public int IncidentID { get; set; }

        /// <summary>
        /// Method which is called when any Workspace Rule Action is invoked.
        /// </summary>
        /// <param name="ActionName">The name of the Workspace Rule Action that was invoked.</param>
        public void RuleActionInvoked(string ActionName)
        {
            try
            {
                if (ActionName == "ValidateHours")
                {
                    int val = 0;
                    int valday = 0;
                    string ata = "";
                    string atatime = "";
                    string atd = "";
                    string atdtime = "";
                    DateTime ATA = DateTime.Now;
                    DateTime ATD = DateTime.Now;
                    Itinerary = _recordContext.GetWorkspaceRecord("CO$Itinerary") as IGenericObject;
                    IList<IGenericField> genericFields;
                    genericFields = Itinerary.GenericFields;
                    foreach (IGenericField gen in genericFields)
                    {
                        if (gen.Name == "ATA")
                        {
                            ata = gen.DataValue.Value.ToString();
                        }
                        if (gen.Name == "ATATime")
                        {
                            atatime = gen.DataValue.Value.ToString();
                        }
                        if (gen.Name == "ATD")
                        {
                            atd = gen.DataValue.Value.ToString();
                        }
                        if (gen.Name == "ATDTime")
                        {
                            atdtime = gen.DataValue.Value.ToString();
                        }
                    }
                    if (!string.IsNullOrEmpty(ata) && !string.IsNullOrEmpty(atatime) && !string.IsNullOrEmpty(ata) && !string.IsNullOrEmpty(atatime))
                    {
                        ATA = DateTime.Parse(ata);
                        ATA = DateTime.Parse(ATA.ToString("yyyy-MM-dd") + " " + atatime);
                        ATD = DateTime.Parse(atd);
                        ATD = DateTime.Parse(ATD.ToString("yyyy-MM-dd") + " " + atdtime);
                        gContext.LogMessage(ATA.ToString() + ":" + ATD.ToString());
                        if (ATD < ATA)
                        {
                            val = 1;
                            //MessageBox.Show("ATA Date is greater than ATD, please check.");

                        }
                        if (ATA < DateTime.Today)
                        {
                            valday = 1;
                            //MessageBox.Show("ATA Date is less than today, please check.");
                        }

                    }
                    foreach (IGenericField gen in genericFields)
                    {
                        if (gen.Name == "ValidateATA_ATD")
                        {

                         

                            gen.DataValue.Value = val;
                        }
                        if (gen.Name == "ValidateATA")
                        {
                          
                          
                            gen.DataValue.Value = valday;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                gContext.LogMessage(e.Message + "Det" + e.StackTrace);
            }
        }

        public string RuleConditionInvoked(string ConditionName)
        {
            return string.Empty;
        }

        #endregion
    }

    [AddIn("Workspace Itinerary AddIn", Version = "1.0.0.0")]
    public class WorkspaceAddInFactory : IWorkspaceComponentFactory2
    {
        public IGlobalContext gContext { get; set; }
        #region IWorkspaceComponentFactory2 Members

        /// <summary>
        /// Method which is invoked by the AddIn framework when the control is created.
        /// </summary>
        /// <param name="inDesignMode">Flag which indicates if the control is being drawn on the Workspace Designer. (Use this flag to determine if code should perform any logic on the workspace record)</param>
        /// <param name="RecordContext">The current workspace record context.</param>
        /// <returns>The control which implements the IWorkspaceComponent2 interface.</returns>
        public IWorkspaceComponent2 CreateControl(bool inDesignMode, IRecordContext RecordContext)
        {
            return new WorkspaceAddIn(inDesignMode, RecordContext, gContext);
        }

        #endregion

        #region IFactoryBase Members

        /// <summary>
        /// The 16x16 pixel icon to represent the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public Image Image16
        {
            get { return Properties.Resources.AddIn16; }
        }

        /// <summary>
        /// The text to represent the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public string Text
        {
            get { return "Itinerary"; }
        }

        /// <summary>
        /// The tooltip displayed when hovering over the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public string Tooltip
        {
            get { return "Itinerary Tooltip"; }
        }

        #endregion

        #region IAddInBase Members

        /// <summary>
        /// Method which is invoked from the Add-In framework and is used to programmatically control whether to load the Add-In.
        /// </summary>
        /// <param name="GlobalContext">The Global Context for the Add-In framework.</param>
        /// <returns>If true the Add-In to be loaded, if false the Add-In will not be loaded.</returns>
        public bool Initialize(IGlobalContext GlobalContext)
        {
            gContext = GlobalContext;
            return true;
        }

        #endregion
    }
}