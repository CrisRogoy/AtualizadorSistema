using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Atualizador_Sistemas.Classes;
using Classes;
using Ionic.Zip;

namespace Atualizador_Sistemas
{
    public partial class FrmPrincipal : Form
    {
        public FrmPrincipal()
        {
            InitializeComponent();
        }

        readonly string sCaminhoSistema = Environment.CurrentDirectory.ToString();
        ZipFile zip = null;
        string sCaminhoZipTemp = "";
        string sNomeArqFor = string.Empty;
        readonly string sTempWindows = Path.GetTempPath();
        string aNomeArquivo = string.Empty;
        readonly Funcoes funcoes = new Funcoes();
        readonly WebClient Wcliente = new WebClient();
        readonly List<string> lListArquivos = new List<string>();
        readonly List<string> lListArquivosTemp = new List<string>();
        readonly List<string> lListCaminhoArquivos = new List<string>();
        int iTotArqDownload = 0;
        int iTotArqJaBaixado = 0;
        bool bProgAberto = true;
        string sPrograma;
        int iErros = 0; // 1 = Erro não localizado arquivo de configuração.
                        // 2 = Erro não localizado arquivo versao.txt.

        [DllImport("wininet.dll")]

        private extern static bool InternetGetConnectedState(out int Description, int ReservedValue);

        public static bool IsConnected()
        {
            return InternetGetConnectedState(out int Description, 0);
        }

        /*--------------------Arrastar Formulario sem bordas--------------------*/
        bool Clicou;
        Point clickedAt;

        private void FrmPrincipal_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;

            Clicou = true;
            clickedAt = e.Location;
        }

        private void FrmPrincipal_MouseMove(object sender, MouseEventArgs e)
        {
            if (Clicou)
            {
                this.Location = new Point(Cursor.Position.X - clickedAt.X, Cursor.Position.Y - clickedAt.Y);
            }
        }

        private void FrmPrincipal_MouseUp(object sender, MouseEventArgs e)
        {
            Clicou = false;
        }
        /*------------------Fim Arrastar Formulario sem bordas------------------*/

        public async Task CheckUpdateAsync(string pReleaseSis)
        {
            lbCancelar.Visible = true;

            string sCaminhoVersao = sCaminhoSistema + "\\versao.txt";

            INIFile ini = new INIFile(sCaminhoVersao);
            string sProdutoVersao = ini.Read("Controle", "Produto");

            if (lListArquivos.Count > 0)
            {
                foreach (string sArquivo in lListArquivos)
                {
                    string sCaminhoPastaTempArq = sTempWindows + sArquivo;

                    if (File.Exists(sCaminhoPastaTempArq + ".zip"))
                    {
                        File.Delete(sCaminhoPastaTempArq + ".zip");
                    }

                    if (Directory.Exists(sCaminhoPastaTempArq))
                    {
                        Directory.Delete(sCaminhoPastaTempArq, true);
                    }

                    lbSair.Visible = false;

                    lListArquivosTemp.Remove(sArquivo);

                    iTotArqJaBaixado++;
                    lbQtdArq.Visible = true;
                    pBox.Visible = true;
                    if (sProdutoVersao == "SIGLIGHT")
                    {
                        await DownloadAsyncUpdate("http://www.teste.com.br/arquivos/" + sArquivo + ".zip", sTempWindows + sArquivo + ".zip");
                    }
                    else if (sProdutoVersao == "SLNFCE")
                    {
                        await DownloadAsyncUpdate("http://www.teste.com.br/arquivos/" + sArquivo + ".zip", sTempWindows + sArquivo + ".zip");
                    }
                    else
                    {
                        await DownloadAsyncUpdate("http://www.teste.com.br/arquivos/" + sArquivo + ".zip", sTempWindows + sArquivo + ".zip");
                    }

                    if (File.Exists(sTempWindows + sArquivo + ".zip"))
                    {
                        ExtrairArquivoZip(sTempWindows + sArquivo + ".zip", sTempWindows);
                        ApagaArquivosEPastas();
                    }

                    for (int i = 0; i < lListArquivos.Count; i++)
                    {
                        string sInicio = sTempWindows + lListArquivos[i].ToString() + "\\" + Path.GetFileName(lListCaminhoArquivos[i]).ToString();
                        string sDestino = sCaminhoSistema + lListCaminhoArquivos[i].ToString();

                        if (File.Exists(sInicio))
                        {
                            CopiaArquivosTempPastaSistema(sInicio, sDestino);
                        }
                    }
                }

                if (File.Exists(sPrograma))
                {
                    if (sPrograma != "" && !string.IsNullOrEmpty(sPrograma))
                    {
                        Process.Start(sPrograma);
                    }
                }
            }
        }

        private void WcArq_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            double JaBaixado = double.Parse((e.BytesReceived / 1024).ToString());
            double Tamanho_Total = double.Parse((e.TotalBytesToReceive / 1024).ToString());
            double Quanto_Falta = Tamanho_Total - JaBaixado;
            //lbStatus.Text = e.ProgressPercentage.ToString() + " % " + "  |  Já baixado : " + JaBaixado + "Kb  |  Tamanho Total : " +
            //                     Tamanho_Total + "Kb  |  Falta : " + Quanto_Falta.ToString("F2") + "Kb";
            pBoxImg.Enabled = false;
            lbStatus.Text = e.ProgressPercentage.ToString() + " %";
            lbStatus.Location = new Point(640 - 320, 182);
            lbStatus.Refresh();

            lbQtdArq.Text = "Total itens para download : " + iTotArqJaBaixado.ToString() + " de " + iTotArqDownload.ToString();
        }

        private void WcArq_DownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
        {
            string sArquivoDelete = sTempWindows + aNomeArquivo + ".zip";

            if (e.Cancelled)
            {
                try
                {
                    Wcliente.Dispose();
                    lListArquivos.Clear();
                    if (File.Exists(sArquivoDelete))
                    {
                        File.Delete(sArquivoDelete);
                    }
                    lbStatus.Location = new Point(640 - 390, 182);
                    lbStatus.Text = "Download Cancelado !";
                    ImgBase64 img = new ImgBase64();
                    pBoxImg.Image = funcoes.ConvertBase64ParaImg(img.ImgNuvemX);
                    pBox.Visible = false;
                    lbSair.Visible = true;
                    lbCancelar.Visible = false;
                }
                catch (Exception Exc)
                {
                    MessageBox.Show(Exc.Message);
                }

                {
                    return;
                }
            }

            if (lListArquivosTemp.Count == 0)
            {
                ImgBase64 img = new ImgBase64();
                pBoxImg.Image = funcoes.ConvertBase64ParaImg(img.ImgSucesso);
                pBoxImg.Location = new Point(260, 50);
                lbCompletoExtrair.Visible = true;
                lbCompletoExtrair.Location = new Point(640 - 390, 205);
                lbCompletoExtrair.Text = "Atualização concluida !";
                lbQtdArq.Visible = false;
                pBox.Visible = false;
                lbCancelar.Visible = false;
                lbSair.Visible = true;
                {
                    return;
                }
            }

            if (!IsConnected())
            {
                Wcliente.Dispose();
                lListArquivos.Clear();
                if (File.Exists(sArquivoDelete))
                {
                    File.Delete(sArquivoDelete);
                }
                lbStatus.Visible = false;
                lbStatus.Location = new Point(640 - 410, 182);
                lbStatus.Text = "Verifique acesso a internet !";
                ImgBase64 img = new ImgBase64();
                pBoxImg.Image = funcoes.ConvertBase64ParaImg(img.ImgNuvemX);
                pBox.Visible = false;
                lbSair.Visible = true;
                lbCancelar.Visible = false;
                lbStatus.Visible = true;

                {
                    return;
                }
            }

            pBox.Visible = false;
        }

        public async Task DownloadAsyncUpdate(string pUrl, string pCaminhoNomeArq)
        {

            Uri uUri = new Uri(pUrl);

            // cria uma instância de webclient

            byte[] bytes = null;
            try
            {
                lbVerAtualizador.Text = "Atualizador versão 1.0";
                Wcliente.DownloadDataCompleted += WcArq_DownloadFileCompleted;
                Wcliente.DownloadProgressChanged += WcArq_DownloadProgressChanged;
                bytes = await Wcliente.DownloadDataTaskAsync(uUri);
            }
            catch (WebException wErro)
            {
                if (wErro.Message == "A solicitação foi anulada: A solicitação foi cancelada.")
                {
                    return;
                }
                else
                {
                    funcoes.GravaTxt("ErroUpdate.log", wErro.ToString(), false);
                    lbSair.Visible = true;
                    lbCancelar.Visible = false;
                    MessageBox.Show(wErro.Message);
                }
            }

            if (bytes != null)
            {
                FileStream Stream = new FileStream(pCaminhoNomeArq, FileMode.Create);

                //Escrevo arquivo no fluxo
                Stream.Write(bytes, 0, bytes.Length);

                //Fecho fluxo para salvar em disco
                Stream.Close();
            }
        }

        private void lbCancelar_MouseClick(object sender, MouseEventArgs e)
        {
            var Msg = MessageBox.Show("Deseja Cancelar download das atualizações ?", "Aviso !", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (Msg == DialogResult.Yes)
            {
                pBox.Visible = false;
                Wcliente.CancelAsync();
                lbCancelar.Visible = false;
                lbSair.Visible = true;
            }
        }

        private void FrmPrincipal_Load(object sender, EventArgs e)
        {
            string sCaminhoVersao = sCaminhoSistema + "\\versao.txt";
            string sCaminhoUpdate = sCaminhoSistema + "\\update.ini";

            if (!File.Exists(sCaminhoVersao))
            {
                ImgBase64 img = new ImgBase64();
                pBoxImg.Image = funcoes.ConvertBase64ParaImg(img.ImgNuvemX);
                pBoxImg.Enabled = false;
                lbStatus.Location = new Point(640 - 430, 182);
                lbStatus.Text = "Não localizado arquivo |Versão.txt|.";
                iErros = 2; // 2 = Erro não localizado arquivo versao.txt.
                lbVerAtualizador.Text = "Atualizador versão 1.0 - Clique para iniciar a atualização.";
                lbCancelar.Visible = false;
                lbSair.Visible = true;
                lbQtdArq.Visible = false;
                pBox.Visible = false;

                {
                    return;
                }
            }

            if (!File.Exists(sCaminhoUpdate))
            {
                ImgBase64 img = new ImgBase64();
                pBoxImg.Image = funcoes.ConvertBase64ParaImg(img.ImgNuvemX);
                pBoxImg.Enabled = false;
                lbStatus.Location = new Point(640 - 490, 182);
                lbStatus.Text = "Não localizado arquivo de configuração |Update.ini|.";
                //lbStatus.Text = "Não localizado arquivo de configuração |Update.ini|.\n                        Clique para configurar.";
                //lbStatus.Cursor = Cursors.Hand;
                iErros = 1; // 1 = Erro não localizado arquivo de configuração.
                lbVerAtualizador.Text = "Atualizador versão 1.0 - Clique para iniciar a atualização.";
                lbCancelar.Visible = false;
                lbSair.Visible = true;
                lbQtdArq.Visible = false;
                pBox.Visible = false;

                {
                    return;
                }
            }

            INIFile ini = new INIFile(sCaminhoVersao);
            string sProdutoVersao = ini.Read("Controle", "Produto");
            string sRelease = ini.Read("Controle", "Release");


            bool bAberto;

            if (sProdutoVersao == "SIGLIGHT")
            {
                sPrograma = Environment.CurrentDirectory.ToString() + "\\SigLight.exe";
                bAberto = FechaProcesso("SigLight");
            }
            else if (sProdutoVersao == "SLNFCE")
            {
                sPrograma = Environment.CurrentDirectory.ToString() + "\\SLNFCe.exe";
                bAberto = FechaProcesso("SLNFCe");
            }
            else
            {
                sPrograma = Environment.CurrentDirectory.ToString() + "\\SLCTe.exe";
                bAberto = FechaProcesso("SLCTe");
            }

            if (bAberto == false)
            {
                {
                    return;
                }
            }

            lListArquivos.Clear();
            lListCaminhoArquivos.Clear();
            lbQtdArq.Visible = false;
            lbCancelar.Visible = false;
            lbCompletoExtrair.Visible = false;
            lbStatus.Text = "";
            lbQtdArq.Text = "";
            pBox.Visible = false;
            lbVerAtualizador.Visible = true;
            lbVerAtualizador.Text = "";

            char[] delimitador = { '+' };
            string[] linhas = { "" };
            char[] delimitadorPipe = { '|' };
            string aRegistro = string.Empty;
            string aDescricao = string.Empty;
            string aCaminhoArquivo = string.Empty;
            string aTamanhoArquivo = string.Empty;
            string aDescricaoCliente = string.Empty;

            if (IsConnected())
            {
                INIFile iArqUpdate = new INIFile(sCaminhoUpdate);
                string UpdateAuto = iArqUpdate.Read("Update", "UpdateAuto");
                string sCNPJEmp = iArqUpdate.Read("Update", "CNPJEmp");
                string sPlanoSis = iArqUpdate.Read("Update", "PLANOSIS");
                string sUpdateAuto = iArqUpdate.Read("Update", "UpdateAuto");

                if (File.Exists(sCaminhoVersao))
                {
                    using (WebClient WcLista = new WebClient())
                    {
                        string sListaArquivos = "";

                        if (sProdutoVersao != "")
                        {
                            if (sProdutoVersao == "SIGLIGHT")
                            {
                                sListaArquivos = WcLista.DownloadString("http://www.teste.com.br/atualizasgl/listaupdate.php?release=" + sRelease + "&pw=4525");
                            }
                            else if (sProdutoVersao == "SLNFCE")
                            {
                                sListaArquivos = WcLista.DownloadString("http://www.teste.com.br/atualizaslnfce/listaupdate.php?release=" + sRelease + "&pw=4525");
                            }
                            else
                            {
                                sListaArquivos = WcLista.DownloadString("http://www.teste.com.br/atualizacte/listaupdate.php?release=" + sRelease + "&pw=4525");
                            }
                            linhas = sListaArquivos.Split(delimitador);
                        }
                        else
                        {
                            ImgBase64 img = new ImgBase64();
                            pBoxImg.Image = funcoes.ConvertBase64ParaImg(img.ImgNuvemX);
                            pBoxImg.Enabled = false;
                            lbStatus.Location = new Point(640 - 410, 170);
                            lbStatus.Text = "Nao localizado 'Produto' !";
                            {
                                return;
                            }
                        }
                    }

                    for (int i = 0; i < linhas.Length - 1; i++)
                    {
                        string[] linhasOk = linhas[i].Split(delimitadorPipe);
                        aRegistro = linhasOk[0]; // Registro;
                        if (aRegistro != "")
                        {
                            aDescricao = linhasOk[1]; // Descrição;
                            aNomeArquivo = linhasOk[2]; // MD5/Nome Arquivo;
                            aCaminhoArquivo = linhasOk[3]; // Caminho Arquivo;
                            aTamanhoArquivo = linhasOk[4]; // Tamanho Arquivo;
                            aDescricaoCliente = linhasOk[5]; // Descrição que cliente vai ver;

                            if (!File.Exists(sCaminhoSistema + aCaminhoArquivo) || funcoes.BytesToString(funcoes.CalculaMD5(sCaminhoSistema + aCaminhoArquivo)).ToUpper() != aNomeArquivo.ToUpper())
                            {
                                lListArquivos.Add(aNomeArquivo);
                                lListCaminhoArquivos.Add(aCaminhoArquivo);
                                lListArquivosTemp.Add(aNomeArquivo);
                            }
                        }
                    }

                    if (lListArquivos.Count == 0 && lListCaminhoArquivos.Count == 0)
                    {
                        ImgBase64 img = new ImgBase64();
                        pBoxImg.Location = new Point(640 - 380, 56);
                        lbCancelar.Visible = false;
                        lbSair.Visible = true;
                        pBoxImg.Image = funcoes.ConvertBase64ParaImg(img.ImgSucesso);
                        pBoxImg.Enabled = false;
                        lbStatus.Visible = true;
                        lbStatus.Location = new Point(640 - 400, 200);
                        lbStatus.Text = "Sistema já atualizado !";
                        lbVerAtualizador.Text = "Atualizador versão 1.0.";
                    }
                    else
                    {
                        if (UpdateAuto == "S")
                        {
                            iTotArqDownload = lListArquivos.Count;
                            ImgBase64 img = new ImgBase64();
                            pBoxImg.Image = funcoes.ConvertBase64ParaImg(img.ImgNuvem);
                            lbStatus.Location = new Point(640 - 430, 182);
                            lbStatus.Text = "";
                            lbVerAtualizador.Text = "";
                            lbSair.Visible = false;
                            lbSair.Visible = true;

                            CheckUpdateAsync(sRelease);
                        }
                        else
                        {
                            iTotArqDownload = lListArquivos.Count;
                            ImgBase64 img = new ImgBase64();
                            pBoxImg.Image = funcoes.ConvertBase64ParaImg(img.ImgNuvem);
                            lbStatus.Location = new Point(640 - 430, 182);
                            lbStatus.Text = "Clique na nuvem para atualizar.";
                            lbVerAtualizador.Text = "Atualizador versão 1.0 - Clique para iniciar a atualização.";
                            lbSair.Visible = false;
                            lbSair.Visible = true;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Verifique acesso a internet e tente novamente.", "Erro !", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        private void lbSair_MouseClick(object sender, MouseEventArgs e)
        {
            Application.Exit();
        }

        private void pBoxImg_MouseClick(object sender, MouseEventArgs e)
        {
            string sCaminhoVersao = sCaminhoSistema + "\\versao.txt";

            pBoxImg.Enabled = false;
            pBoxImg.Refresh();

            if (File.Exists(sCaminhoVersao))
            {
                INIFile iArqVersao = new INIFile(sCaminhoVersao);
                string sSVN = iArqVersao.Read("Controle", "SVN");
                string sVersao = iArqVersao.Read("Controle", "versao");
                string sRelease = iArqVersao.Read("Controle", "Release");
                string sData_Compilacao = iArqVersao.Read("Controle", "Data_Compilacao");

                CheckUpdateAsync(sRelease);
            }
        }

        private void pBoxImg_MouseMove(object sender, MouseEventArgs e)
        {
            ImgBase64 img = new ImgBase64();

            pBoxImg.Image = funcoes.ConvertBase64ParaImg(img.ImgNuvemClara);
        }

        private void pBoxImg_MouseLeave(object sender, EventArgs e)
        {
            ImgBase64 img = new ImgBase64();

            pBoxImg.Image = funcoes.ConvertBase64ParaImg(img.ImgNuvem);
        }

        public void ExtrairArquivoZip(string pLocalZip, string pDestino)
        {
            if (File.Exists(pLocalZip))
            {
                if (Directory.Exists(pDestino))
                {
                    try
                    {
                        zip = new ZipFile(pLocalZip);
                        zip.ExtractProgress += Zip_ExtractProgress;
                        sCaminhoZipTemp = pLocalZip;
                        zip.ExtractAll(pDestino + Directory.CreateDirectory(Path.GetFileNameWithoutExtension(pLocalZip)));
                    }
                    catch (Exception Err)
                    {
                        MessageBox.Show(Err.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Local destino não definido.");
                }
            }
            else
            {
                MessageBox.Show("O arquivo Zip nao foi localizado.");
            }
        }

        private void Zip_ExtractProgress(object sender, ExtractProgressEventArgs e)
        {
            double dTotalBytesParaTransferir = e.TotalBytesToTransfer;
            double dTotalBytesTransferidos = e.BytesTransferred;
            double dFalta = ((dTotalBytesParaTransferir - dTotalBytesTransferidos) / 1024);
            lbCompletoExtrair.Text = "Extraindo dados : " + dFalta.ToString();
            lbCompletoExtrair.Refresh();

            if (dTotalBytesTransferidos > 0 && dTotalBytesTransferidos == dTotalBytesParaTransferir)
            {
                zip.Dispose();
                File.Delete(sCaminhoZipTemp);

                foreach (string sNomArq in lListCaminhoArquivos)
                {
                    string sNomeArquivo = sNomArq;
                    int iLenArq = sNomeArquivo.Length;
                    string sArquivoRenomeado = sNomeArquivo.Insert(iLenArq - 2, "_").Insert(iLenArq, "_").Insert
                                               (iLenArq + 2, "_" + DateTime.Now.ToString("ddMMyyyyhhmmss"));

                    if (File.Exists(sCaminhoSistema + sNomeArquivo))
                    {
                        funcoes.RenomeiaArquivo(sCaminhoSistema + sNomeArquivo, sCaminhoSistema + sArquivoRenomeado);
                    }
                }
            }
        }

        private void ApagaArquivosEPastas()
        {
            foreach (string sArquivos in lListArquivos)
            {
                string sNomeArquivoPasta = Path.GetFileName(sArquivos);

                if (Directory.Exists(sCaminhoSistema + "\\" + sNomeArquivoPasta))
                {
                    Directory.Delete(sCaminhoSistema + "\\" + sNomeArquivoPasta, true);
                }
            }
            lbCompletoExtrair.Location = new Point(640 - 440, 205);
            lbCompletoExtrair.Text = "Atualização concluida com sucesso !";
            lbCompletoExtrair.Refresh();
        }

        private void CopiaArquivosTempPastaSistema(string pInicio, string pDestino)
        {
            FileStream fsDestino = new FileStream(pDestino, FileMode.Create);
            FileStream fsInicio = new FileStream(pInicio, FileMode.Open);
            byte[] bt = new byte[1048756];
            int readByte;

            while ((readByte = fsInicio.Read(bt, 0, bt.Length)) > 0)// && !ParaCopia)
            {
                fsDestino.Write(bt, 0, readByte);
                double dPerc = (fsInicio.Position * 100 / fsInicio.Length);
                //lbStatus.Visible = true;
                lbStatus.Text = "Copiando : " + dPerc.ToString() + " %";
                lbStatus.Location = new Point(640 - 385, 170);
                lbStatus.Refresh();
            }

            fsDestino.Close();
            fsInicio.Close();
        }

        private void lbSite_MouseClick(object sender, MouseEventArgs e)
        {
            var Msg = MessageBox.Show("Clique em SIM para ser direcionado ao site da teste.", "Abrir site ? ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (Msg == DialogResult.Yes)
            {
                Process.Start("http://teste.com.br/");
            }
        }

        private bool FechaProcesso(string sProcesso)
        {
            Process[] pProcesso = Process.GetProcessesByName(sProcesso);

            int iQtdprocessos = 0;
            if (pProcesso.Length > 0)
            {
                var Msg = MessageBox.Show("Existem processos impedindo a atualização do sistema :  \n| " + sProcesso + " | Deseja fechar o processo ?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (Msg == DialogResult.Yes)
                {
                    while (bProgAberto)
                    {
                        foreach (Process p in pProcesso)
                        {
                            iQtdprocessos++;
                            p.Kill();
                        }
                        bProgAberto = false;
                    }
                    return true; // Fechados Todos
                }
                else
                {
                    ImgBase64 img = new ImgBase64();
                    pBoxImg.Image = funcoes.ConvertBase64ParaImg(img.ImgNuvemX);
                    pBoxImg.Enabled = false;
                    lbStatus.Location = new Point(640 - 550, 182);
                    lbStatus.Text = "Existe um ou mais processos impedindo a execução da atualização.";
                    lbVerAtualizador.Text = "Atualizador versão 1.0 - Clique para iniciar a atualização.";
                    lbCancelar.Visible = false;
                    lbSair.Visible = true;
                    lbQtdArq.Visible = false;
                    pBox.Visible = false;
                    return false; // Aberto Todos
                }
            }
            else
            {
                return true;
            }
        }

        private void lbStatus_Click(object sender, EventArgs e)
        {
            if (iErros == 1)
            {
                MessageBox.Show(iErros.ToString());
            }
        }
    }
}