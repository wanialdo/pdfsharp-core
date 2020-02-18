using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using estagioprobatorioback.Context;
using estagioprobatorioback.Models;
using estagioprobatorioback.Reports;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using System.Xml.XPath;
using System.IO;

namespace estagioprobatorioback.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class RelatorioFinalController : ControllerBase
    {
        private readonly EPContext _context;
        //Document document { get; set; }

        public RelatorioFinalController(EPContext context)
        {
            _context = context;
        }

        // GET: api/Avaliacao
        [HttpGet, Route("{nr_matricula}")]
        public FileStreamResult GetAvaliacao(int nr_matricula)
        {
            //Buscar no Banco
            Avaliacao_Final model = GenerateModel(nr_matricula);

            ReportAvaliacaoFinal report = new ReportAvaliacaoFinal(model);
            return report.CreateDocument();
        }

        protected Avaliacao_Final GenerateModel(int nr_matricula)
        {
            var avaliacoesServidor = _context.GetDataFromDatabase();
            var servidorConsiste = _context.GetDataFromDatabase();

            Avaliacao_Final avaliacao = new Avaliacao_Final();
            avaliacao.Nome = servidorConsiste.nome;
            avaliacao.Cargo = servidorConsiste.nome_cargo;
            avaliacao.Matricula = servidorConsiste.nr_matricula;
            avaliacao.Filial = servidorConsiste.org_sigla;
            avaliacao.Setor = servidorConsiste.descricaolot;
            avaliacao.DataAdmissao = servidorConsiste.dtadmissao;

            foreach (var avaliacaoServidor in avaliacoesServidor)
            {
                avaliacao.ResultadosAvaliacao.Add(new ResultadoAvaliacaoFinal()
                {
                    NomeEtapa = avaliacaoServidor.Periodo_Avaliacao.Etapa.dsc_etapa,
                    DataInicioEtapa = avaliacaoServidor.Periodo_Avaliacao.dat_inicio.ToString("dd/MM/yyyy"),
                    DataFimEtapa = avaliacaoServidor.Periodo_Avaliacao.dat_fim.ToString("dd/MM/yyyy"),
                    ResultadoEtapa = avaliacaoServidor.resultadoAvaliacao.ToString()
                });
            }

            avaliacao.ResultadoFinal = "0";
            avaliacao.Conclusao = true;

            avaliacao.MembrosComissao.Add(new MembroComissaoAvaliacao() { NomeMembroComissao = " ", MatriculaMembroComissao = " " });
            avaliacao.MembrosComissao.Add(new MembroComissaoAvaliacao() { NomeMembroComissao = " ", MatriculaMembroComissao = " " });
            avaliacao.MembrosComissao.Add(new MembroComissaoAvaliacao() { NomeMembroComissao = " ", MatriculaMembroComissao = " " });
            avaliacao.MembrosComissao.Add(new MembroComissaoAvaliacao() { NomeMembroComissao = " ", MatriculaMembroComissao = " " });
            avaliacao.MembrosComissao.Add(new MembroComissaoAvaliacao() { NomeMembroComissao = " ", MatriculaMembroComissao = " " });
            avaliacao.MembrosComissao.Add(new MembroComissaoAvaliacao() { NomeMembroComissao = " ", MatriculaMembroComissao = " " });

            return avaliacao;
        }
    }
}
