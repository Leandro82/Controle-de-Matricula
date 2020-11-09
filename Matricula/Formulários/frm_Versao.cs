using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Deployment.Application;

namespace Matricula
{
    public partial class frm_Versao : Form
    {
        Boolean doUpdate;
        public frm_Versao(Boolean vr)
        {
            InitializeComponent();
            doUpdate = vr;
        }

        private void button1_Click(object sender, EventArgs e)
        {
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
                    MessageBox.Show("A nova versão do aplicativo não pode ser baixada no momento. \n\nVerifique sua conexão de rede ou tente novamente mais tarde. Error: " + dde.Message);
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

                if (doUpdate)
                {
                    try
                    {
                        ad.Update();
                        MessageBox.Show("O aplicativo foi atualizado e agora será reiniciado.");
                        Application.Restart();
                    }
                    catch (DeploymentDownloadException dde)
                    {
                        MessageBox.Show("Não é possível instalar a versão mais recente do aplicativo. \n\nPor favor, verifique sua conexão de rede ou tente novamente mais tarde. Error: " + dde);
                        return;
                    }
                }
            }
        }
    }
}
