using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace estagioprobatorioback.Models
{
    public class Avaliacao_Final
    {
        public Avaliacao_Final() {
            ResultadosAvaliacao = new List<ResultadoAvaliacaoFinal>();
            MembrosComissao = new List<MembroComissaoAvaliacao>();
        }

        public String Nome { get; set; }
        public String Cargo { get; set; }
        public String Matricula { get; set; }
        public String Filial { get; set; }
        public String Setor { get; set; }
        public String DataAdmissao { get; set; }
        public IList<ResultadoAvaliacaoFinal> ResultadosAvaliacao { get; set; }
        public String ResultadoFinal { get; set; }
        public bool Conclusao { get; set; }
        public IList<MembroComissaoAvaliacao> MembrosComissao { get; set; }
    }   

    public class ResultadoAvaliacaoFinal 
    {
        public String NomeEtapa { get; set; }
        public String DataInicioEtapa { get; set; }
        public String DataFimEtapa { get; set; }
        public String ResultadoEtapa { get; set; }
    }

    public class MembroComissaoAvaliacao 
    {
        public String NomeMembroComissao { get; set; }
        public String MatriculaMembroComissao { get; set; }
    }
}