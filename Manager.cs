using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Linq;
using Automate.Utils;
using Automate.Models;
using Automate.Services.Applications;
using Automate.Exceptions;
using System.Web.Script.Serialization;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Diagnostics;

namespace Automate
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Manager
    {
        public static ManagerParameters parameters;
        private int excelPId;

        private enum Field
        {
            Action = 1,
            Target,
            Value,
            Wait,
            Return,
            Label
        }
        
        public Manager(string[] args)
        {
            parameters = this.CreateParams();
            parameters.Deserialize(args);
            SetConsoleCtrlHandler(new TypeHandler(ConsoleCtrlHandler), true);
        }
       
        public string Run()
        {
            ManagerReturn result = this.Run(parameters.Get("$robotPath$").ToString(), parameters.Get("$Caminho Download Final$").ToString(), parameters);
            string output = result.Serialize(parameters.Get("$outputType$").ToString().ToLower());

            return StringUtils.EncodeString(output);
        }

        public void SetParameters(string[] args)
        {
            parameters = this.CreateParams();
            parameters.Deserialize(args);
        }
        #region consoleCtrlHandler

        public delegate bool TypeHandler(CtrlTypes CtrlType);

        [DllImport("Kernel32")]
        public static extern bool SetConsoleCtrlHandler(TypeHandler Handler, bool Add);

        public enum CtrlTypes
        {
            CTRL_C_EVENT = 0,
            CTRL_BREAK_EVENT,
            CTRL_CLOSE_EVENT,
            CTRL_LOGOFF_EVENT = 5,
            CTRL_SHUTDOWN_EVENT
        }

        private bool ConsoleCtrlHandler(CtrlTypes ctrlType)
        {
            Log.Debug("Console event triggered");

            switch (ctrlType)
            {
                case CtrlTypes.CTRL_C_EVENT:
                case CtrlTypes.CTRL_BREAK_EVENT:
                case CtrlTypes.CTRL_CLOSE_EVENT:
                case CtrlTypes.CTRL_LOGOFF_EVENT:
                case CtrlTypes.CTRL_SHUTDOWN_EVENT:
                    CloseExcel();
                    break;
                default:
                    Log.Debug("Unrecognized event type: " + ctrlType.ToString());
                    break;
            }
            return true;
        }

        #endregion

        [ComVisible(true)]
        public ManagerParameters CreateParams()
        {
            return new ManagerParameters();
        }

        [ComVisible(true)]
        public ManagerReturn CreateReturn()
        {
            return new ManagerReturn();
        }

        [ComVisible(true)]
        private ManagerReturn Run(string robotPath, string downloadPath, ManagerParameters myParameters)
        {
            Range xlRange;
            ManagerApps apps = new ManagerApps();
            ManagerReturn managerReturn = CreateReturn();
            Application oXL;

            var pids = GetPIdsByName("excel");
            oXL = new Application();
            excelPId = GetPIdsByName("excel").Except(pids).ElementAt(0);

            try
            {
                xlRange = OpenXlsx(oXL, new FileInfo(robotPath));                
            }
            catch (Exception e)
            {
                LogException(managerReturn, e, "Ocorreu um erro ao abrir o robô especificado");
                return managerReturn;
            }

            int i = 0;
            try
            {
                for (i = 2; ; i++)
                {
                    if (xlRange.Cells[i, Field.Action] == null || xlRange.Cells[i, Field.Action].Value2 == null)
                    {
                        managerReturn.Add("result", "1");
                        managerReturn.Add("errorMessage", "");
                        break;
                    }
                    
                    DoAction(xlRange, ref i, downloadPath, apps, myParameters, managerReturn);
                }
            }
            catch (Exception e)
            {
                LogException(managerReturn, e, "Exception thrown at row: " + i);
                if(Convert.ToString(myParameters.Get("$ifExceptionCloseApps$")) == "1")
                {
                    apps.Close();
                }
            }
            finally
            {
                foreach (Workbook wb in oXL.Workbooks)
                {
                    wb.Close(false);
                }

                oXL.DisplayAlerts = false;
                oXL.Workbooks.Close();
                Marshal.FinalReleaseComObject(oXL);
                oXL = null;

                CloseExcel();
            }

            return managerReturn;
        }

        private void DoAction(Range xlRange, ref int i, string downloadPath, ManagerApps apps, ManagerParameters myParameters, ManagerReturn managerReturn)
        {
            JavaScriptSerializer serializer;
            string action = xlRange.Cells[i, Field.Action].Value2;
            string target = GetTarget(xlRange, i, myParameters, managerReturn);
            string value  = GetValue(xlRange, i, myParameters, managerReturn);
            int    wait   = Convert.ToInt32(xlRange.Cells[i, Field.Wait].Value2);
            string returnVar = xlRange.Cells[i, Field.Return].Value2;
            object returnValue;

            if (action.Contains("(HIDDEN)"))
            {
                action = action.Replace("(HIDDEN)", "").Trim();
            }
            else
            {
                Log.Debug(string.Join(" | ", new string[] { action, target, value, wait.ToString(), returnVar }));
            }


            switch (action)
            {
                case "Abrir": // Temporary for compatibility with old robots
                case "AbrirBrowser":
                    // TODO: Receive what browser with target parameter
                    if (target == "IE" || target == "Internet Explorer")
                    {
                        returnValue = apps.OpenApplication<InternetExplorer>(value).ToString();
                    } else if (target =="PhantomJS") {
                        returnValue = apps.OpenApplication<PhantomJS>(value).ToString();
                    } else {
                        returnValue = apps.OpenApplication<Chrome>(value).ToString();
                    }
                    AddToReturn(returnValue.ToString(), returnVar, managerReturn);
                    break;
                case "AbrirDesktopApp":
                    returnValue = apps.OpenApplication<DesktopApp>(target).ToString();
                    AddToReturn(returnValue.ToString(), returnVar, managerReturn);
                    break;
                case "Preencher":
                    apps.activeApp.Fill(target, value, wait);
                    break;
                case "Limpar":
                    apps.activeApp.Clear(target, wait);
                    break;
                case "EsperarElemento":
                    apps.activeApp.WaitShow(target, wait);
                    break;
                case "Esperar":
                    SO.Wait(wait);
                    break;
                case "Clicar":
                    apps.activeApp.Click(target, wait, value);
                    break;
                case "ClicarPorMouse":
                    apps.activeApp.ClickByMouse(target, wait, value);
                    break;
                case "Baixar":
                    returnValue = apps.activeApp.Download(target, GetDownloadPath(value), wait);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "BaixarURL":
                    returnValue = DownloadURL(target, GetDownloadPath(value), wait);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "Captcha":
                    apps.activeApp.Captcha(target, value, wait);
                    break;  
                case "IrPara":
                    returnValue = JumpToIf(true, ref i, GetRowIndex(xlRange, value));
                    AddToReturn(returnValue, returnVar, managerReturn);
                    SO.Wait(wait);
                    break;
                case "SeElementoExiste":
                    JumpToIf(apps.activeApp.ElementExist(target, wait), ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeNaoElementoExisteId":
                    JumpToIf(!apps.activeApp.ElementExistId(target, wait), ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeElementoExisteId":
                    JumpToIf(apps.activeApp.ElementExistId(target, wait), ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeNaoElementoExisteClass":
                    JumpToIf(!apps.activeApp.ElementExistClass(target, wait), ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeElementoExisteClass":
                    JumpToIf(apps.activeApp.ElementExistClass(target, wait), ref i, GetRowIndex(xlRange, value));
                    break;
                case "EsperarAlerta":
                case "EsperarAlertaOuElemento":
                    apps.activeApp.WaitAlertOrElement(target, wait);
                    break;
                case "SeAlertaExiste":
                case "SeAlertaOuElementoExiste":
                    JumpToIf(apps.activeApp.AlertOrElementExist(target, wait), ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeNaoAlertaExiste":
                case "SeNaoAlertaOuElementoExiste":
                    JumpToIf(!apps.activeApp.AlertOrElementExist(target, wait), ref i, GetRowIndex(xlRange, value));
                    break;
                case "LerTextoAlerta":
                    returnValue = apps.activeApp.GetAlertText(wait);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "ClicarAlerta":
                    apps.activeApp.ClickAlert(target, value, wait);
                    break;
                case "EsperarPaginaCarregar":
                    apps.activeApp.WaitPageLoad(wait);
                    break;
                case "Navegar":
                    apps.activeApp.Navigate(value, wait);
                    break;
                case "NavegarParaElemento":
                    apps.activeApp.NavigateToElement(target, wait);
                    break;
                case "Imprimir":
                    returnValue = apps.activeApp.Print(target, GetDownloadPath(value), wait);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "LerXPS":
                    returnValue = XPS.Read(target);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "BuscarNoTexto":
                    returnValue = Text.SearchFor(target, value);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "LerTexto":
                    returnValue = apps.activeApp.GetText(target, wait);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "LerTextoPorClique":
                    returnValue = apps.activeApp.GetTextByClick(target, wait);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "LerTodosTextos":
                    returnValue = apps.activeApp.GetAllTexts(target, wait);
                    serializer = new JavaScriptSerializer();
                    returnValue = serializer.Serialize(returnValue);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "LerTodosAtributos":
                    returnValue = apps.activeApp.GetAllAtributes(target, value, wait);
                    serializer = new JavaScriptSerializer();
                    returnValue = serializer.Serialize(returnValue);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "RemoverTexto":
                    returnValue = Text.Remove(target, value);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "LerTabela":
                    returnValue = apps.activeApp.GetTable(target, wait);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "CriaPasta":                    
                    returnValue = SO.CreateFolder(value);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "EscreverArquivo":
                    SO.WriteToFile(target, value);
                    break;
                case "FecharAplicacao":
                    apps.Close(int.Parse(target));
                    break;
                case "FecharTodasAsAplicacoes":
                    apps.Close();
                    AddToReturn(value, returnVar, managerReturn);
                    break;
                case "Teclado":
                    SO.TypeKeyboard(value);
                    SO.Wait(wait);
                    break;
                case "SelecionarItem":
                    apps.activeApp.SelectFromDropdown(target, value, wait);
                    break;
                case "SeNaoElementoExiste":
                    JumpToIf(!apps.activeApp.ElementExist(target, wait), ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeVariavelVazia":
                    JumpToIf((target.Trim().Equals("")), ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeVariavelIgual":
                    JumpToIf(CompareTextWithArray(target.Split(';')[0],
                                                    target.Split(';').Skip(1).ToArray(),
                                                    (str1, str2) => str1 == str2),
                            ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeNaoVariavelIgual":
                    JumpToIf(!CompareTextWithArray(target.Split(';')[0],
                                                    target.Split(';').Skip(1).ToArray(),
                                                    (str1, str2) => str1 == str2),
                            ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeVariavelContem":
                    JumpToIf(CompareTextWithArray(target.Split(';')[0],
                                                    target.Split(';').Skip(1).ToArray(),
                                                    (str1, str2) => str1.Contains(str2)),
                            ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeNaoVariavelContem":
                    JumpToIf(!CompareTextWithArray(target.Split(';')[0],
                                                    target.Split(';').Skip(1).ToArray(),
                                                    (str1, str2) => str1.Contains(str2)),
                            ref i, GetRowIndex(xlRange, value));
                    break;
                case "IrParaProximaAba":
                    apps.activeApp.SwitchToTab(1);
                    break;
                case "IrParaAbaPrincipal":
                    apps.activeApp.SwitchToTab(0);
                    break;
                case "SeElementoNaoEstaVazio":
                    JumpToIf(!apps.activeApp.ElementIsEmpty(target, wait), ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeElementoEstaVazio":
                    JumpToIf(apps.activeApp.ElementIsEmpty(target, wait), ref i, GetRowIndex(xlRange, value));
                    break;
                case "SelecionarIFrame":
                    apps.activeApp.SwitchToFrame(target, value, wait);
                    break;
                case "SelecionarMainFrame":
                    apps.activeApp.SwitchToMainFrame(wait);
                    break;
                case "RodarJavaScript":
                case "RodarJavascript":
                    returnValue = apps.activeApp.RunJavaScript(value);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "LerDropDown":
                    returnValue = apps.activeApp.GetDropdownOptions(target, value, wait);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "LerLista":
                    returnValue = apps.activeApp.GetListOptions(target, value, wait);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "AceitarAlerta":
                    apps.activeApp.AcceptAlert(wait);
                    break;
                case "SalvarArquivo":
                    returnValue = apps.activeApp.SaveFile(GetDownloadPath(value), wait);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "LerAtributo":
                    returnValue = apps.activeApp.GetAttribute(target, value, wait);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "AtribuirVariavel":
                    AddToReturn(value, returnVar, managerReturn);
                    break;
                case "IncrementarVariavel":
                    AddToReturn((Convert.ToInt32(target) + Convert.ToInt32(value)).ToString(), returnVar, managerReturn);
                    break;                
                case "SeTextoIgual":
                    JumpToIf(CompareTextWithArray(apps.activeApp.GetText(target.Split(';')[0], wait),
                                                    target.Split(';').Skip(1).ToArray(),
                                                    (str1, str2) => str1 == str2),
                            ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeNaoTextoIgual":
                    JumpToIf(!CompareTextWithArray(apps.activeApp.GetText(target.Split(';')[0], wait),
                                                    target.Split(';').Skip(1).ToArray(),
                                                    (str1, str2) => str1 == str2),
                            ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeTextoContem":
                    JumpToIf(CompareTextWithArray(apps.activeApp.GetText(target.Split(';')[0], wait),
                                                    target.Split(';').Skip(1).ToArray(),
                                                    (str1, str2) => str1.Contains(str2)),
                            ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeNaoTextoContem":
                    JumpToIf(!CompareTextWithArray(apps.activeApp.GetText(target.Split(';')[0], wait),
                                                    target.Split(';').Skip(1).ToArray(),
                                                    (str1, str2) => str1.Contains(str2)),
                            ref i, GetRowIndex(xlRange, value));
                    break;
                case "Renomear":
                    returnValue = SO.MoveFile(target, value);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "SubstituirEmString":
                    returnValue = target.Replace(value.Split(';')[0], value.Split(';')[1]);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "Log":
                    LogMessage(target, value);
                    break;
                case "Screenshot":
                    apps.activeApp.TakeScreenshot(GetDownloadPath(value));
                    break;
                case "PrepararDownload":
                    apps.activeApp.DownloadSetUp();
                    break;
                case "EsperarDownload":
                    returnValue = apps.activeApp.WaitDownload(wait);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "FinalizarDownload":
                    returnValue = apps.activeApp.DownloadCleanUp(target, GetDownloadPath(value));
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "SeDownloadAtivo":
                    JumpToIf(apps.activeApp.OngoingDownload(wait), ref i, GetRowIndex(xlRange, value));
                    break;
                case "SeNaoDownloadAtivo":
                    JumpToIf(!apps.activeApp.OngoingDownload(wait), ref i, GetRowIndex(xlRange, value));
                    break;
                case "EsconderBrowser":
                    apps.activeApp.HideBrowser();
                    break;
                case "MostrarBrowser":
                    apps.activeApp.ShowBrowser();
                    break;
                case "MostrarAlerta":
                    System.Windows.Forms.MessageBox.Show(value, "Alerta!", System.Windows.Forms.MessageBoxButtons.OK);
                    break;
                case "ClicarPorTexto":
                    apps.activeApp.ClickByText(target, wait);
                    break;
                case "RemoverRepeticoes":
                    returnValue = RemoveChars(target, value);
                    AddToReturn(returnValue, returnVar, managerReturn);
                    break;
                case "Upload":
                    apps.activeApp.Upload(target, value, wait);
                    break;
                case "MudarFocoBrowser":
                    apps.activeApp.ChangeWindowFocus(value);
                    break;
                default:
                    Log.Debug(action + " não existe meu jovem, favor modificá-la");
                    break;
            }
        }

        private string RemoveChars(string str, string sequence)
        {
            foreach (char c in sequence)
            {
                str = String.Join(c.ToString(), str.Split(new char[] {c}, StringSplitOptions.RemoveEmptyEntries));
            }

            return str;
        }

        private List<int> GetPIdsByName(string name)
        {
            return (from process in Process.GetProcessesByName(name) select process.Id).ToList();
        }

        private string GetDownloadPath(string value)
        {
            string[] path = value.Split('\\');

            if (path.Length > 1)
            {
                return value;
            }

            return Path.Combine(Manager.parameters.Get("$Caminho Download Final$").ToString(), value);
        }

        private void LogMessage(string file, string message)
        {
            if (file == "")
            {
                Log.Debug(message);
                return;
            }

            using (FileStream fs = File.Open(file, FileMode.OpenOrCreate))
            {
                fs.Seek(0, SeekOrigin.End);
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    sw.WriteLine(message);
                    sw.Close();
                }
                fs.Close();
            }

        }

        private string GetRowIndex(Range xlRange, string label)
        {
            if (Regex.IsMatch(label, @"^\d+$"))
            {
                return label;
            }

            for (int i = 2; i <= xlRange.Count; i++)
            {
                if (xlRange.Cells[i, Field.Label].value2 == label)
                {
                    return i.ToString();
                }
            }

            throw new NonExistentLabelException(label);
        }

        private string DownloadURL(string url, string destFile, int timeout)
        {
            WebClientTimeout client = new WebClientTimeout(timeout);

            try
            {
                client.DownloadFile(new Uri(url), destFile);
            }
            catch (System.Net.WebException e)
            {
                return "Timeout: " + e.ToString();
            }
            catch (System.Exception e)
            {
                return "Failed: " + e.GetType();
            }

            return destFile;
        }

        private string JumpToIf(Boolean condition, ref int i, string value)
        {
            string nextNaturalI = (i + 1).ToString();

            if(condition)
            {
                i = Convert.ToInt32(value) - 1;
            }

            return nextNaturalI;
        }

        private void AddToReturn(object returnValue, string returnVar, ManagerReturn managerReturn)
        {
            if (returnVar != null)
            {
                managerReturn.Add(returnVar, Convert.ToString(returnValue).Trim('\r').Trim('\n'));
            }
        }

        private string GetTarget(Range xlRange, int i, ManagerParameters myParameters, ManagerReturn managerReturn)
        {
            return GetValueFromParamsAndReturn(xlRange, i, Field.Target, myParameters, managerReturn);
        }

        private string GetValue(Range xlRange, int i, ManagerParameters myParameters, ManagerReturn managerReturn)
        {
            return GetValueFromParamsAndReturn(xlRange, i, Field.Value, myParameters, managerReturn);
        }

        private string GetValueFromParamsAndReturn(Range xlRange, int row, Field field, ManagerParameters myParameters, ManagerReturn managerReturn)
        {
            string xlsValue = Convert.ToString(xlRange.Cells[row, field].Text);

            if (xlsValue == null) return null;
                       
            foreach (var key in managerReturn.Keys.Concat(myParameters.Keys))
            {
                try
                {
                    string keyValue = managerReturn.ContainsKey(key) ? (string) managerReturn.Get(key) : (string) myParameters.Get(key);
                    xlsValue = xlsValue.Replace(key, keyValue);
                } catch { }
            }

            return xlsValue;
        }

        private static Range OpenXlsx(Application oXL, FileInfo file)
        {
            Workbook xlWorkbook = oXL.Workbooks.Open(file.FullName);
            _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;
            return xlRange;
        }

        private void CloseExcel()
        {
            try
            {
                Process.GetProcessById(excelPId).Kill();
            }
            catch
            {
                Log.Debug("Failed to kill Excell process");
            }
        }

        private Boolean CompareTextWithArray(string text, string[] array, Func<string, string, Boolean> compare)
        {
            foreach (string str in array)
            {
                if (compare(text.Trim(), str.Trim()))
                {
                    return true;
                }
            }

            return false;
        }

        private void LogException(ManagerReturn managerReturn, Exception e, string extraMessage = "")
        {
            managerReturn.Add("result", "0");
            managerReturn.Add("errorMessage", e.Message);
            managerReturn.Add("stackTrace", e.StackTrace);
            managerReturn.Add("errorExtraMessage", extraMessage);
            Log.Debug(DateTime.Now + " - " + extraMessage + ": " + e.Message + ". Detailed Error: " + e.StackTrace);
            SO.WriteToFile("./manager.log", DateTime.Now + " - " + extraMessage + ": " + e.Message + ". Detailed Error: " + e.StackTrace);
        }
    }
}
