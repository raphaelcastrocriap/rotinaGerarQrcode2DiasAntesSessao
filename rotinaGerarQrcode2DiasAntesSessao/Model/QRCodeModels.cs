using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace RotinaGerarQrcode2DiasAntesSessao.Model
{
    /// <summary>
    /// Dados de uma sessão com formador vindos da query SQL.
    /// </summary>
    public class SessaoFormador
    {
        public int VersaoRowid { get; set; }
        public DateTime? Data { get; set; }
        public string HoraInicio { get; set; }
        public string HoraFim { get; set; }
        public int? RowidModulo { get; set; }
        public string NumSessao { get; set; }
        public string NomeAbreviado { get; set; }
        public string Descricao { get; set; }
        public int NumeroAccao { get; set; }
        public string RefAccao { get; set; }
        public int CodigoFormador { get; set; }
        public string Email { get; set; }
    }

    /// <summary>
    /// Request para a API de geração do QRCode F029B.
    /// </summary>
    public class QRCodeApiRequest
    {
        [JsonProperty("refAcao")]
        public string RefAcao { get; set; }

        [JsonProperty("rowIdsSessoes")]
        public List<int> RowIdsSessoes { get; set; }
    }

    /// <summary>
    /// Resposta da API de geração do QRCode F029B.
    /// </summary>
    public class QRCodeApiResponse
    {
        [JsonProperty("ambiente")]
        public string Ambiente { get; set; }

        [JsonProperty("sucesso")]
        public bool Sucesso { get; set; }

        [JsonProperty("mensagem")]
        public string Mensagem { get; set; }

        [JsonProperty("totalProcessado")]
        public int TotalProcessado { get; set; }

        [JsonProperty("totalSucesso")]
        public int TotalSucesso { get; set; }

        [JsonProperty("totalFalhas")]
        public int TotalFalhas { get; set; }

        [JsonProperty("sessoes")]
        public List<QRCodeSessaoResult> Sessoes { get; set; } = new List<QRCodeSessaoResult>();
    }

    /// <summary>
    /// Resultado individual de cada sessão gerada pela API.
    /// </summary>
    public class QRCodeSessaoResult
    {
        [JsonProperty("rowIdSessao")]
        public int RowIdSessao { get; set; }

        [JsonProperty("numeroSessao")]
        public string NumeroSessao { get; set; }

        [JsonProperty("dataSessao")]
        public string DataSessao { get; set; }

        [JsonProperty("pathDocx")]
        public string PathDocx { get; set; }

        [JsonProperty("pathPdf")]
        public string PathPdf { get; set; }

        [JsonProperty("sucesso")]
        public bool Sucesso { get; set; }

        [JsonProperty("mensagemErro")]
        public string MensagemErro { get; set; }
    }

    /// <summary>
    /// Item do relatório final de execução da rotina.
    /// </summary>
    public class RelatorioItem
    {
        public string RefAccao { get; set; }
        public string Descricao { get; set; }
        public string NomeFormador { get; set; }
        public string EmailFormador { get; set; }
        public string NumSessao { get; set; }
        public string DataSessao { get; set; }

        /// <summary>
        /// "OK", "ERRO_API", "ERRO_EMAIL", "ERRO_GERACAO"
        /// </summary>
        public string Status { get; set; }
        public string Mensagem { get; set; }
    }
}
