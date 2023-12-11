using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UDFWarning_PopUp.Common
{
    class clsMenuEvent
    {     

        SAPbouiCOM.Form objform;
        string strsql;
        public void MenuEvent_For_StandardMenu(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (clsModule. objaddon.objapplication.Forms.ActiveForm.TypeEx)
                {
                    case "-392":
                    case "-393":
                    case "392":
                    case "393":
                        {
                            // Default_Sample_MenuEvent(pVal, BubbleEvent)
                            if (pVal.BeforeAction == true)
                                return;
                            objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                            Default_Sample_MenuEvent(pVal, BubbleEvent);

                            break;
                        }
                    //case "REVREC":
                    //    RevenueRecognition_MenuEvent(ref pVal, ref BubbleEvent);
                    //    break;
                   
                }
            } 
            catch (Exception ex)
            {

            }
        }

        private void Default_Sample_MenuEvent(SAPbouiCOM.MenuEvent pval, bool BubbleEvent)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                if (pval.BeforeAction == true)
                {
                }

                else
                {
                    SAPbouiCOM.Form oUDFForm;
                    try
                    {
                        oUDFForm = clsModule.objaddon.objapplication.Forms.Item(objform.UDFFormUID);
                    }
                    catch (Exception ex)
                    {
                        oUDFForm = objform;
                    }

                    switch (pval.MenuUID)
                    {
                        case "1281": // Find                            
                              
                                break;                            
                        case "1287":                            
                               
                                break;
                            
                        default:                           
                               
                                break;
                            
                    }
                }
            }
            catch (Exception ex)
            {
                // objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            }
        }

        private void RevenueRecognition_MenuEvent(ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
        {
            try
            {
                SAPbobsCOM.Recordset objRs;
                SAPbouiCOM.DBDataSource DBSource;
                SAPbouiCOM.Matrix Matrix0;
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                DBSource = objform.DataSources.DBDataSources.Item("@AT_REV_RECO");
                Matrix0 = (SAPbouiCOM.Matrix)objform.Items.Item("mtxcont").Specific;
                if (pval.BeforeAction == true)
                {
                    switch (pval.MenuUID)
                    {
                        case "1284": //Cancel
                            if (clsModule.objaddon.objapplication.MessageBox("Cancelling of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") != 1)
                            {
                                BubbleEvent = false;
                            }
                            else
                            {
                                if (((SAPbouiCOM.EditText)objform.Items.Item("tvochid").Specific).String != "")
                                {
                                    objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    //if( Remove_JournalVoucher(objform.UniqueID, Convert.ToInt32(((SAPbouiCOM.EditText)objform.Items.Item("tvochid").Specific).String)) == false) 
                                    if (clsModule.objaddon.HANA)
                                    {
                                        strsql = "Update \"@AT_REV_RECO\" Set \"U_VoucherID\"=null Where \"DocEntry\"=" + objform.DataSources.DBDataSources.Item("@AT_REV_RECO").GetValue("DocEntry", 0) + " ";
                                    }
                                    else
                                    {
                                        strsql = "Update @AT_REV_RECO Set U_VoucherID=null Where DocEntry=" + objform.DataSources.DBDataSources.Item("@AT_REV_RECO").GetValue("DocEntry", 0) + " ";
                                    }
                                    objRs.DoQuery(strsql);
                                    //BubbleEvent = false;
                                }
                            }
                            break;
                        case "1286":
                            {
                                //clsModule.objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                //BubbleEvent = false;
                                //return;
                                break;
                            }
                        case "1293":
                            if (Matrix0.VisualRowCount == 1) BubbleEvent = false;
                            break;
                    }
                }
                else
                {                    
                    switch (pval.MenuUID)
                    {
                        case "1281": // Find Mode                            
                                objform.Items.Item("tdocnum").Enabled = true;
                                objform.Items.Item("tglc").Enabled = true;
                                objform.Items.Item("tgln").Enabled = true;
                                objform.Items.Item("series").Enabled = true;
                                objform.Items.Item("tposdate").Enabled = true;                            
                                objform.Items.Item("mtxcont").Enabled = false;
                            objform.Items.Item("tvochid").Enabled = true;
                            objform.Items.Item("tjeid").Enabled = true;
                            objform.EnableMenu("1282", true);
                            objform.ActiveItem = "tdocnum";
                            break;
                            
                        case "1282"://Add Mode                            
                                clsModule.objaddon.objglobalmethods.LoadSeries(objform, DBSource, "AT_REVREC");
                                ((SAPbouiCOM.EditText)objform.Items.Item("tposdate").Specific).String = "A";//DateTime.Now.Date.ToString("dd/MM/yy");
                                ((SAPbouiCOM.EditText)objform.Items.Item("trem").Specific).String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss");
                            ((SAPbouiCOM.ComboBox)objform.Items.Item("cprojtype").Specific).Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                //clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "invnum", "#");
                            objform.EnableMenu("1282", false);
                            break;
                        case "1293"://Delete Row
                            //DeleteRow(Matrix0, "@REV_RECO1");
                            break;
                                              
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }       

        public void DeleteRow(SAPbouiCOM.Matrix objMatrix, string TableName)
        {
            try
            {
                SAPbouiCOM.DBDataSource DBSource;
                // objMatrix = objform.Items.Item("20").Specific
                objMatrix.FlushToDataSource();
                DBSource = objform.DataSources.DBDataSources.Item(TableName); 
                for (int i = 1, loopTo = objMatrix.VisualRowCount; i <= loopTo; i++)
                {
                    objMatrix.GetLineData(i);
                    DBSource.Offset = i - 1;
                    DBSource.SetValue("LineId", DBSource.Offset, Convert.ToString(i));
                    objMatrix.SetLineData(i);
                    objMatrix.FlushToDataSource();
                }
                DBSource.RemoveRecord(DBSource.Size - 1);
                objMatrix.LoadFromDataSource();
            }

            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
            }
            finally
            {
            }
        }        

    }
}
