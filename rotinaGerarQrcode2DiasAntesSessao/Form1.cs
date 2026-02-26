using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;
using RotinaGerarQrcode2DiasAntesSessao.Connects;
using RotinaGerarQrcode2DiasAntesSessao.Model;
using RotinaGerarQrcode2DiasAntesSessao.Properties;
using static RotinaGerarQrcode2DiasAntesSessao.Model.GestorEmailConfirmadosModel;

namespace RotinaGerarQrcode2DiasAntesSessao
{
    public partial class Form1 : Form
    {
        // ── Configurações ────────────────────────────────────────────────────────
        //private const string API_URL              = "http://localhost:5141/api/v2/acoes-dtp/gerar-f029b-qrcode"; // teste
        private const string API_URL              = "http://192.168.1.213:8080/api/v2/acoes-dtp/gerar-f029b-qrcode";
        private const string CC_EMAIL_PEDAGOGICO  = "tecnicopedagogico@criap.com";
        private const string EMAIL_INFORMATICA    = "informatica@criap.com";
        private const string EMAIL_TESTE          = "raphaelcastro@criap.com";

        // Formadores excluídos (coordenadores / pessoal interno)
        private static readonly int[] FORMADORES_EXCLUIDOS = { 699, 704, 827, 1046, 1053, 15683, 15684, 1425, 16221 };

        // ── Estado da execução ────────────────────────────────────────────────────
        public bool teste = false;
        public string versao;
        private List<RelatorioItem> relatorio = new List<RelatorioItem>();

        // ─────────────────────────────────────────────────────────────────────────
        public Form1()
        {
            InitializeComponent();
            Security.remote();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Version v = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text += " V." + v.Major + "." + v.Minor + "." + v.Build;
            versao = $@" <br><font size=""-2"">Versão: V.{v.Major}.{v.Minor}.{v.Build}"
                   + $@" | Build: {File.GetLastWriteTime(Assembly.GetExecutingAssembly().Location)}"
                   + @" | rotinaGerarQrcode2DiasAntesSessao</font>";

            string[] args = Environment.GetCommandLineArgs();
            if (args.Contains("-a") || args.Contains("-A"))
            {
                try
                {
                    ExecutarRotinaQRCode();
                }
                catch (Exception ex)
                {
                    EnviarEmailErro(ex.ToString());
                }
                finally
                {
                    Application.Exit();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            richTextBox1.Clear();
            try
            {
                ExecutarRotinaQRCode();
                Log("Rotina concluída.");
            }
            catch (Exception ex)
            {
                Log("ERRO: " + ex.Message);
                MessageBox.Show("Erro: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                EnviarEmailErro(ex.ToString());
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        // ══════════════════════════════════════════════════════════════════════════
        //  ROTINA PRINCIPAL  (síncrona)
        // ══════════════════════════════════════════════════════════════════════════
        private void ExecutarRotinaQRCode()
        {
            relatorio.Clear();

            // Sessões 2 dias à frente (ex.: hoje=23/02 → procura sessões de 25/02)
            DateTime targetDate = DateTime.Today.AddDays(2);
            Log("Buscando sessões presenciais para " + targetDate.ToString("dd/MM/yyyy") + "...");

            // ── 1. Consultar sessões ──────────────────────────────────────────────
            List<SessaoFormador> sessoes = QuerySessoes(targetDate);
            Log(sessoes.Count + " registo(s) encontrado(s).");

            if (sessoes.Count == 0)
            {
                EnviarEmailRelatorio(targetDate,
                    "Nenhuma sessão presencial (sala) encontrada para " + targetDate.ToString("dd/MM/yyyy") + ".");
                return;
            }

            // ── 2. Agrupar por RefAccao e chamar a API ────────────────────────────
            var grupos = sessoes.GroupBy(s => s.RefAccao);

            using (HttpClient httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(60) })
            {
                foreach (var grupo in grupos)
                {
                    string refAccao = grupo.Key;
                    List<int> rowIds = grupo.Select(s => s.VersaoRowid).Distinct().ToList();

                    Log("-> Gerando QRCode para '" + refAccao + "' (" + rowIds.Count + " sessão(ões))...");

                    // ── 3. Chamar API ─────────────────────────────────────────────
                    QRCodeApiResponse apiResponse = GerarQRCodeViaApi(httpClient, refAccao, rowIds);

                    if (apiResponse == null || !apiResponse.Sucesso)
                    {
                        string errMsg = (apiResponse != null ? apiResponse.Mensagem : null) ?? "Sem resposta da API";
                        Log("  ERRO API - " + refAccao + ": " + errMsg);

                        relatorio.Add(new RelatorioItem
                        {
                            RefAccao = refAccao,
                            Status   = "ERRO_API",
                            Mensagem = errMsg
                        });
                        continue;
                    }

                    Log("  API OK: " + apiResponse.TotalSucesso + " gerada(s), " + apiResponse.TotalFalhas + " falha(s).");

                    // ── 4. Registar falhas de geração no relatório ────────────────
                    foreach (var sessaoErro in apiResponse.Sessoes.Where(s => !s.Sucesso))
                    {
                        Log("  FALHA GERACAO sessao " + sessaoErro.NumeroSessao + ": " + sessaoErro.MensagemErro);
                        relatorio.Add(new RelatorioItem
                        {
                            RefAccao  = refAccao,
                            NumSessao = sessaoErro.NumeroSessao,
                            DataSessao= sessaoErro.DataSessao,
                            Status    = "ERRO_GERACAO",
                            Mensagem  = sessaoErro.MensagemErro ?? "Falha ao gerar QRCode"
                        });
                    }

                    // ── 5. Para cada sessão gerada com sucesso → enviar email ─────
                    foreach (var sessaoResult in apiResponse.Sessoes.Where(s => s.Sucesso))
                    {
                        // Todos os formadores associados a esta sessão
                        List<SessaoFormador> formadoresDaSessao = sessoes
                            .Where(s => s.VersaoRowid == sessaoResult.RowIdSessao)
                            .ToList();

                        foreach (var formador in formadoresDaSessao)
                        {
                            try
                            {
                                EnviarEmailFormador(formador, sessaoResult);
                                Log("  OK Email -> " + formador.NomeAbreviado + " (" + formador.Email + ") | Sessao " + sessaoResult.NumeroSessao);

                                RegistraLog(
                                    formador.CodigoFormador.ToString(),
                                    "QRCode F029B gerado e email enviado | Sessao " + sessaoResult.NumeroSessao
                                        + " | " + sessaoResult.DataSessao
                                        + " | PDF: " + sessaoResult.PathPdf,
                                    "QRCode Gerado - F029B",
                                    refAccao
                                );

                                relatorio.Add(new RelatorioItem
                                {
                                    RefAccao      = refAccao,
                                    Descricao     = formador.Descricao,
                                    NomeFormador  = formador.NomeAbreviado,
                                    EmailFormador = formador.Email,
                                    NumSessao     = sessaoResult.NumeroSessao,
                                    DataSessao    = sessaoResult.DataSessao,
                                    Status        = "OK",
                                    Mensagem      = "Email enviado com sucesso"
                                });
                            }
                            catch (Exception ex)
                            {
                                Log("  ERRO EMAIL -> " + formador.NomeAbreviado + ": " + ex.Message);
                                relatorio.Add(new RelatorioItem
                                {
                                    RefAccao      = refAccao,
                                    Descricao     = formador.Descricao,
                                    NomeFormador  = formador.NomeAbreviado,
                                    EmailFormador = formador.Email,
                                    NumSessao     = sessaoResult.NumeroSessao,
                                    DataSessao    = sessaoResult.DataSessao,
                                    Status        = "ERRO_EMAIL",
                                    Mensagem      = ex.Message
                                });
                            }
                        }
                    }
                }
            }

            // ── 6. Relatório final ────────────────────────────────────────────────
            Log("Enviando relatorio final...");
            EnviarEmailRelatorio(targetDate);
            Log("Rotina finalizada.");
        }

        // ══════════════════════════════════════════════════════════════════════════
        //  QUERY SQL — sessões presenciais com 2 dias de antecedência
        // ══════════════════════════════════════════════════════════════════════════
        private List<SessaoFormador> QuerySessoes(DateTime targetDate)
        {
            string dataFiltro = targetDate.ToString("yyyy-MM-dd");
            string excluidos  = string.Join(",", FORMADORES_EXCLUIDOS);

            string query = @"
                SELECT DISTINCT
                    s.versao_rowid,
                    s.Data,
                    s.Hora_Inicio,
                    s.Hora_Fim,
                    s.Rowid_Modulo,
                    s.Num_Sessao,
                    f.Nome_Abreviado,
                    cu.Descricao,
                    a.Numero_Accao,
                    a.Ref_Accao,
                    f.Codigo_Formador,
                    COALESCE(c.Email1, c.Email2) AS Email
                FROM TBForSessoesFormadores sf
                INNER JOIN TBForSessoes s      ON s.versao_rowid = sf.rowid_sessao
                INNER JOIN TBForAccoes a        ON s.Rowid_Accao = a.versao_rowid
                INNER JOIN TBForFormadores f    ON f.Codigo_Formador = sf.codigo_formador
                INNER JOIN TBGerContactos c     ON f.versao_rowid = c.Codigo_Entidade
                                               AND c.Tipo_Entidade = 4
                INNER JOIN TBForCursos cu       ON cu.Codigo_Curso = a.Codigo_Curso
                WHERE CAST(s.Data AS DATE) = '" + dataFiltro + @"'
                  AND s.Comp_elr = 'P'
                  AND f.Codigo_Formador NOT IN (" + excluidos + @")
                  AND COALESCE(c.Email1, c.Email2) IS NOT NULL";

            var result = new List<SessaoFormador>();

            Connect.HTlocalConnect.ConnInit();
            try
            {
                DataTable dt = new DataTable();
                using (SqlDataAdapter adapter = new SqlDataAdapter(query, Connect.HTlocalConnect.Conn))
                {
                    adapter.Fill(dt);
                }

                foreach (DataRow row in dt.Rows)
                {
                    result.Add(new SessaoFormador
                    {
                        VersaoRowid    = Convert.ToInt32(row["versao_rowid"]),
                        Data           = row.IsNull("Data")          ? (DateTime?)null : (DateTime)row["Data"],
                        HoraInicio     = row.IsNull("Hora_Inicio")   ? null : row["Hora_Inicio"].ToString(),
                        HoraFim        = row.IsNull("Hora_Fim")      ? null : row["Hora_Fim"].ToString(),
                        RowidModulo    = row.IsNull("Rowid_Modulo")  ? (int?)null : Convert.ToInt32(row["Rowid_Modulo"]),
                        NumSessao      = row.IsNull("Num_Sessao")    ? null : row["Num_Sessao"].ToString(),
                        NomeAbreviado  = row.IsNull("Nome_Abreviado")? null : row["Nome_Abreviado"].ToString().Trim(),
                        Descricao      = row.IsNull("Descricao")     ? null : row["Descricao"].ToString().Trim(),
                        NumeroAccao    = row.IsNull("Numero_Accao")  ? 0     : Convert.ToInt32(row["Numero_Accao"]),
                        RefAccao       = row.IsNull("Ref_Accao")     ? null : row["Ref_Accao"].ToString().Trim(),
                        CodigoFormador = row.IsNull("Codigo_Formador") ? 0   : Convert.ToInt32(row["Codigo_Formador"]),
                        Email          = row.IsNull("Email")         ? null : row["Email"].ToString().Trim()
                    });
                }
            }
            finally
            {
                Connect.HTlocalConnect.ConnEnd();
            }

            return result;
        }

        // ══════════════════════════════════════════════════════════════════════════
        //  CHAMADA À API — POST gerar-f029b-qrcode
        // ══════════════════════════════════════════════════════════════════════════
        private QRCodeApiResponse GerarQRCodeViaApi(HttpClient httpClient, string refAccao, List<int> rowIds)
        {
            try
            {
                var requestBody = new QRCodeApiRequest
                {
                    RefAcao       = refAccao,
                    RowIdsSessoes = rowIds
                };

                string json = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                HttpResponseMessage httpResponse = httpClient.PostAsync(API_URL, content).Result;
                string responseBody = httpResponse.Content.ReadAsStringAsync().Result;

                if (httpResponse.IsSuccessStatusCode)
                    return JsonConvert.DeserializeObject<QRCodeApiResponse>(responseBody);

                return new QRCodeApiResponse
                {
                    Sucesso  = false,
                    Mensagem = "HTTP " + (int)httpResponse.StatusCode + ": " + responseBody
                };
            }
            catch (Exception ex)
            {
                return new QRCodeApiResponse
                {
                    Sucesso  = false,
                    Mensagem = "Excecao ao chamar API: " + ex.Message
                };
            }
        }

        // ══════════════════════════════════════════════════════════════════════════
        //  EMAIL PARA O FORMADOR  (mensagem HTML + PDF em anexo)
        // ══════════════════════════════════════════════════════════════════════════
        private void EnviarEmailFormador(SessaoFormador formador, QRCodeSessaoResult sessaoResult)
        {
            string dataFormatada = formador.Data.HasValue
                ? formador.Data.Value.ToString("dd/MM/yyyy")
                : sessaoResult.DataSessao;

            string subject = $"Instituto CRIAP || Folha Sum\u00e1rio QRCode - {formador.Descricao} - {dataFormatada}";

            // Formatar horas como "19h00"
            string horaInicio = FormatarHora(formador.HoraInicio);
            string horaFim    = FormatarHora(formador.HoraFim);

            // Data no formato dd.MM.yyyy (conforme modelo institucional)
            string dataFormatadaPonto = formador.Data.HasValue
                ? formador.Data.Value.ToString("dd.MM.yyyy")
                : dataFormatada.Replace("/", ".");

            var sb = new StringBuilder();
            sb.AppendLine($"<p>Estimado(a) Professor(a) <b>{formador.NomeAbreviado}</b>,</p>");
            sb.AppendLine("<p>Fazemos votos de que se encontre bem.</p>");
            sb.AppendLine($"<p>No seguimento da aula prevista para o dia <b>{dataFormatadaPonto}</b>, a decorrer no hor&aacute;rio das <b>{horaInicio}</b> &agrave;s <b>{horaFim}</b>, procedemos ao envio, em anexo ao presente email, da(s) folha(s) de presen&ccedil;as com o respetivo QR Code.</p>");
            sb.AppendLine("<p>Informamos igualmente que o referido documento j&aacute; se encontra dispon&iacute;vel na plataforma Moodle, na sec&ccedil;&atilde;o denominada <b>&ldquo;Registo QR Code / Assiduidade&rdquo;</b>.</p>");
            sb.AppendLine("<p>Para aceder &agrave; plataforma, dever&aacute; utilizar as credenciais de acesso, anteriormente enviadas por e-mail.</p>");
            sb.AppendLine("<p>Para projetar o documento no in&iacute;cio da sess&atilde;o de forma&ccedil;&atilde;o, por forma a permitir o registo da sua assiduidade e dos formandos, poder&aacute; utilizar a via (e-mail ou plataforma Moodle) que considerar mais adequada.</p>");
            sb.AppendLine("<p>Refor&ccedil;amos que, idealmente, a folha de presen&ccedil;as dever&aacute; ser projetada <b>10 minutos antes</b> da hora de in&iacute;cio da sess&atilde;o de forma&ccedil;&atilde;o.</p>");
            sb.AppendLine("<p>Permanecemos ao dispor para qualquer esclarecimento adicional.</p>");
            sb.AppendLine("<p>Com os melhores cumprimentos,<br><b>Departamento T&eacute;cnico-Pedag&oacute;gico</b><br>Instituto CRIAP</p>");

            using (SmtpClient client = CriarSmtpClient())
            using (MailMessage mm = new MailMessage())
            {
                mm.From = new MailAddress("Instituto CRIAP <" + Settings.Default.emailenvio + ">");

                if (!teste)
                {
                    mm.To.Add(formador.Email);
                    mm.CC.Add(CC_EMAIL_PEDAGOGICO);
                    mm.CC.Add(EMAIL_INFORMATICA);
                    mm.ReplyToList.Add(new MailAddress(CC_EMAIL_PEDAGOGICO));
                }
                else
                {
                    mm.To.Add(EMAIL_TESTE);
                }

                mm.Subject             = subject;
                mm.BodyEncoding        = Encoding.UTF8;
                mm.IsBodyHtml          = true;
                mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
                // Rodapé com versão nunca vai para o formador
                mm.Body = CriarHtmlEmail("Folha de Sum&aacute;rio QRCode (F029B)", sb.ToString(), rodapeInterno: false);

                // Anexar o PDF se o ficheiro existir no caminho de rede
                if (!string.IsNullOrWhiteSpace(sessaoResult.PathPdf) && File.Exists(sessaoResult.PathPdf))
                    mm.Attachments.Add(new Attachment(sessaoResult.PathPdf));

                client.Send(mm);
            }
        }

        // ══════════════════════════════════════════════════════════════════════════
        //  EMAIL DE RELATÓRIO FINAL
        // ══════════════════════════════════════════════════════════════════════════
        private void EnviarEmailRelatorio(DateTime targetDate, string mensagemExtra = null)
        {
            try
            {
                string dataFormatada = targetDate.ToString("dd/MM/yyyy");

                int okCount    = relatorio.Count(r => r.Status == "OK");
                int erroCount = relatorio.Count(r => r.Status != "OK");

                var linhas = new StringBuilder();
                foreach (var item in relatorio)
                {
                    string rowBg = item.Status == "OK" ? "#fff" : "#ffe4d6";
                    linhas.AppendLine($@"<tr style='background:{rowBg};'>
                        <td>{item.RefAccao}</td>
                        <td>{item.Descricao}</td>
                        <td>{item.NomeFormador}</td>
                        <td>{item.EmailFormador}</td>
                        <td style='text-align:center;'>{item.NumSessao}</td>
                        <td>{item.DataSessao}</td>
                        <td style='text-align:center;'><b>{item.Status}</b></td>
                        <td>{item.Mensagem}</td>
                      </tr>");
                }

                string tabelaHtml = relatorio.Count > 0
                    ? $@"<p style='margin:12px 0;'>
                        <b>Total:</b> {relatorio.Count} &nbsp;|&nbsp;
                        <b style='color:#27ae60;'>Sucesso:</b> {okCount} &nbsp;|&nbsp;
                        <b style='color:#c0392b;'>Erros:</b> {erroCount}
                       </p>
                       <table border='0' cellpadding='5' cellspacing='0'
                              style='border-collapse:collapse;font-size:12px;width:100%;'>
                         <tr style='background:#ed7520;color:#fff;'>
                           <th>Ref A&ccedil;&atilde;o</th><th>Curso</th><th>Formador</th>
                           <th>Email</th><th>Sess&atilde;o N&ordm;</th><th>Data/Hora</th>
                           <th>Status</th><th>Mensagem</th>
                         </tr>
                         {linhas}
                       </table>"
                    : !string.IsNullOrEmpty(mensagemExtra)
                        ? $"<p>{mensagemExtra}</p>"
                        : "<p>Nenhum item processado.</p>";

                using (SmtpClient client = CriarSmtpClient())
                using (MailMessage mm = new MailMessage())
                {
                    mm.From = new MailAddress("Instituto CRIAP <" + Settings.Default.emailenvio + ">");

                    if (!teste)
                    {
                        mm.To.Add(EMAIL_INFORMATICA);
                        mm.To.Add(CC_EMAIL_PEDAGOGICO);
                    }
                    else
                    {
                        mm.To.Add(EMAIL_TESTE);
                    }

                    mm.Subject      = $"Instituto CRIAP || Relat\u00f3rio QRCode F029B - Sess\u00f5es {dataFormatada}";
                    mm.BodyEncoding = Encoding.UTF8;
                    mm.IsBodyHtml   = true;
                    mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
                    mm.Body = CriarHtmlEmail($"Relat&oacute;rio QRCode F029B &mdash; Sess&otilde;es de {dataFormatada}", tabelaHtml, rodapeInterno: true) + "<br>" + versao;

                    client.Send(mm);
                }
            }
            catch (Exception ex)
            {
                EnviarEmailErro("Falha ao enviar relatorio final: " + ex.ToString());
            }
        }

        // ══════════════════════════════════════════════════════════════════════════
        //  EMAIL DE ERRO
        // ══════════════════════════════════════════════════════════════════════════
        private void EnviarEmailErro(string detalhe)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("<p style='color:#c0392b;'><b>Ocorreu um erro na rotina rotinaGerarQrcode2DiasAntesSessao.</b></p>");
                sb.AppendLine("<pre style='background:#ffe4d6;border:1px solid #ed7520;padding:12px;font-size:11px;white-space:pre-wrap;word-break:break-all;'>");
                sb.AppendLine(detalhe);
                sb.AppendLine("</pre>");

                using (SmtpClient client = CriarSmtpClient())
                using (MailMessage mm = new MailMessage())
                {
                    mm.From    = new MailAddress("Instituto CRIAP <" + Settings.Default.emailenvio + ">");
                    mm.To.Add(!teste ? EMAIL_INFORMATICA : EMAIL_TESTE);
                    mm.Subject      = "ERRO - rotinaGerarQrcode2DiasAntesSessao";
                    mm.BodyEncoding = Encoding.UTF8;
                    mm.IsBodyHtml   = true;
                    mm.Body = CriarHtmlEmail("Erro na Rotina QRCode F029B", sb.ToString(), rodapeInterno: true) + "<br>" + versao;
                    client.Send(mm);
                }
            }
            catch { /* silently fail on error-email */ }
        }

        // ══════════════════════════════════════════════════════════════════════════
        //  LOG BASE DE DADOS  (secretariaVirtual.sv_logs)
        // ══════════════════════════════════════════════════════════════════════════
        public void RegistraLog(string id, string mensagem, string menu, string refAcao)
        {
            DataBaseLogSave(new List<objLogSendersFormador>
            {
                new objLogSendersFormador
                {
                    idFormador = id,
                    mensagem   = (mensagem ?? "").Replace("'", "''"),
                    menu       = menu,
                    refAccao   = refAcao
                }
            });
        }

        public static void DataBaseLogSave(List<objLogSendersFormador> logSenders)
        {
            if (logSenders == null || logSenders.Count == 0) return;

            var sb = new StringBuilder(
                "INSERT INTO sv_logs (idFormando, refAcao, dataregisto, registo, menu, username) VALUES ");

            for (int i = 0; i < logSenders.Count; i++)
            {
                string sep = i < logSenders.Count - 1 ? "," : "";
                sb.Append("('"
                    + logSenders[i].idFormador + "', '"
                    + logSenders[i].refAccao   + "', GETDATE(), '"
                    + logSenders[i].mensagem   + "', '"
                    + logSenders[i].menu       + "', 'system_rotina')" + sep);
            }

            Connect.SVlocalConnect.ConnInit();
            try
            {
                using (SqlCommand cmd = new SqlCommand(sb.ToString(), Connect.SVlocalConnect.Conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            finally
            {
                Connect.SVlocalConnect.ConnEnd();
            }
        }

        // ══════════════════════════════════════════════════════════════════════════
        //  HELPERS
        // ══════════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Gera o HTML completo com o layout padrão CRIAP (#ed7520).
        /// <param name="rodapeInterno">Se true, acrescenta rodapé com "Instituto CRIAP — envio automático" (apenas emails internos).</param>
        /// </summary>
        private static string CriarHtmlEmail(string titulo, string conteudo, bool rodapeInterno = false)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html><html>");
            sb.AppendLine("<head><meta charset='utf-8'>");
            sb.AppendLine("<style>");
            sb.AppendLine("  body  { font-family: Arial, sans-serif; font-size: 13px; color: #333; margin: 0; padding: 0; }");
            sb.AppendLine("  .header { padding: 16px 24px; }");
            sb.AppendLine("  .header h2 { color: #333; margin: 0; font-size: 15px; font-weight: bold; }");
            sb.AppendLine("  .content { padding: 20px 24px; }");
            sb.AppendLine("  table { border-collapse: collapse; width: auto; }");
            sb.AppendLine("  table th { background: #ed7520; color: #fff; padding: 7px 12px; text-align: left; border: 1px solid #d4641a; }");
            sb.AppendLine("  table td { padding: 6px 12px; border: 1px solid #eee; }");
            sb.AppendLine("  .footer { font-size: 11px; color: #999; padding: 10px 24px 16px; border-top: 2px solid #ed7520; margin-top: 20px; }");
            sb.AppendLine("</style></head>");
            sb.AppendLine("<body>");
            sb.AppendLine("  <div class='content'>");
            sb.AppendLine(conteudo);
            sb.AppendLine("  </div>");
            if (rodapeInterno)
                sb.AppendLine("  <div class='footer'>Instituto CRIAP &mdash; envio autom&aacute;tico</div>");
            sb.AppendLine("</body></html>");
            return sb.ToString();
        }

        private SmtpClient CriarSmtpClient()
        {
            return new SmtpClient
            {
                Port           = 25,
                Host           = "mail.criap.com",
                Timeout        = 10000,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                Credentials    = new NetworkCredential(
                    Settings.Default.emailenvio,
                    Settings.Default.passwordemail)
            };
        }

        /// <summary>
        /// Converte "19:00:00", "19:00" ou "26/02/2026 13:30:00" em "13h30".
        /// </summary>
        private static string FormatarHora(string hora)
        {
            if (string.IsNullOrWhiteSpace(hora)) return "";
            // Tenta como DateTime completo (ex: "26/02/2026 13:30:00")
            if (DateTime.TryParse(hora, out DateTime dt))
                return dt.ToString("HH") + "h" + dt.ToString("mm");
            // Tenta como TimeSpan puro (ex: "13:30:00" ou "13:30")
            if (TimeSpan.TryParse(hora, out TimeSpan ts))
                return ts.Hours.ToString("D2") + "h" + ts.Minutes.ToString("D2");
            return hora;
        }

        private void Log(string msg)
        {
            string line = "[" + DateTime.Now.ToString("HH:mm:ss") + "] " + msg;
            Console.WriteLine(line);

            if (richTextBox1 != null && !richTextBox1.IsDisposed)
            {
                if (richTextBox1.InvokeRequired)
                    richTextBox1.Invoke((MethodInvoker)(() => {
                        richTextBox1.AppendText(line + Environment.NewLine);
                        richTextBox1.ScrollToCaret();
                    }));
                else
                {
                    richTextBox1.AppendText(line + Environment.NewLine);
                    richTextBox1.ScrollToCaret();
                }
            }
        }
    }
}
