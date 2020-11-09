using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Deployment.Application;
using System.Diagnostics;

namespace Matricula
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static bool IsAppAlreadyRunning()
        {
            Process currentProcess = Process.GetCurrentProcess();
            if (Process.GetProcessesByName(currentProcess.ProcessName).Any(p => p.Id != currentProcess.Id && !p.HasExited))
            {
                return true;
            }

            return false;
        }

        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            UpdateCheckInfo info = null;

            if (ApplicationDeployment.IsNetworkDeployed)
            {
                ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;

                try
                {
                    info = ad.CheckForDetailedUpdate();
                }
                catch (DeploymentDownloadException dde)
                {
                    MessageBox.Show("A nova versão do aplicativo não pode ser baixada no momento. \n\nVerifique sua conexão de rede ou tente novamente mais tarde. Error:" + dde.Message);
                    return;
                }
                catch (InvalidDeploymentException ide)
                {
                    MessageBox.Show("Não é possível procurar por uma nova versão do aplicativo. A implantação do ClickOnce está corrompida. Por favor, reimplemente o aplicativo e tente novamente. Error: " + ide.Message);
                    return;
                }
                catch (InvalidOperationException ioe)
                {
                    MessageBox.Show("Este aplicativo não pode ser atualizado. Provavelmente não é um aplicativo ClickOnce. Error: " + ioe.Message);
                    return;
                }

                if (info.UpdateAvailable)
                {
                    Boolean doUpdate = true;
                    if (!info.IsUpdateRequired && IsAppAlreadyRunning() == false)
                    {
                        var peq = new frm_Versao(doUpdate);
                        if (Application.OpenForms.OfType<frm_Versao>().Count() > 0)
                        {
                            Application.OpenForms[peq.Name].Focus();
                        }
                        else
                        {
                            peq.ShowDialog();
                        }
                    }
                }
            }
            if (IsAppAlreadyRunning() == false)
            {
                Application.Run(new frm_Principal());
            }
            else
            {
                string msg = "O Sistema de RM já está aberto";
                frm_Mensagem mg = new frm_Mensagem(msg);
                mg.ShowDialog();
            }
        }
    }
}
