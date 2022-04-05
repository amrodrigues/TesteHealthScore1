using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using Projeto.Entities;

namespace TesteHealthScore
{
    class Program
    {
        static void Main(string[] args)
        {

            List<Events> listaevents = new List<Events>();

            var xls = new XLWorkbook(@"C:\Users\amrod\Downloads\Teste Técnico CSharp\Model.xlsx");

            var planilha = xls.Worksheet("Events");
            var planilha2 = xls.Worksheet("Settings");
            var planilha3 = xls.Worksheet("Outputs");


            var totalLinhas = planilha.Rows().Count();
     
            //Tabela Events
            for (int l = 3; l <= totalLinhas; l++)
            {
                Events e = new Events();
                e.Id = int.Parse(planilha.Cell($"A{l}").Value.ToString());
                e.Name = planilha.Cell($"B{l}").Value.ToString();
                e.HealthScoreDiscount = int.Parse(planilha.Cell($"C{l}").Value.ToString());
                listaevents.Add(e);
            }


            //tabelaSettings
            Settings s = new Settings();
            s.BaseHealthScore = int.Parse(planilha2.Cell($"B3").Value.ToString());
            s.Profiles = int.Parse(planilha2.Cell($"B2").Value.ToString());

            Console.WriteLine(s.BaseHealthScore.ToString());
            Console.WriteLine(s.Profiles.ToString());
            Console.WriteLine("===============");


            //tabela Outputs
            List<Outputs> listaoutput = new List<Outputs>();
            for (int contador = 1; contador <= s.Profiles; contador++)
            {
                Outputs o = new Outputs();
                o.ProfileId = contador;
                Random rand = new Random();
                int nrale = rand.Next(1, 10);
                o.HealthScore = s.BaseHealthScore + nrale;
                listaoutput.Add(o);
            }

            Random rand2 = new Random();


            string nome_cabecalho = "";
            int valor_helth = 1;

            List<Outputs> listavirtal = new List<Outputs>();

          
            for (int l = 3; l <= 13; l++)
            {
                
                planilha3.Cell($"A{l}").Value = valor_helth;

            }


                // while ((nome_cabecalho != "Death") && (valor_helth != 0))
                for (int contador = 1; contador <= s.Profiles; contador++)
            {
                int nrale2 = rand2.Next(1, 5);

                foreach (Events p in listaevents)
                {
                    if (p.Id == nrale2)
                    {
                        nome_cabecalho = p.Name;
                        valor_helth = p.HealthScoreDiscount;

                        //    planilha3.Cell($"{letra}2").Value = nome_cabecalho;


                        for (int contadorx = 1; contadorx <= s.Profiles; contadorx++)
                        {

                            Outputs o = new Outputs();
                            o.ProfileId = contador;
                            o.HealthScore = s.BaseHealthScore - valor_helth;
                            listavirtal.Add(o);

                            Console.WriteLine(o.HealthScore.ToString());
                        }


                        Console.WriteLine("===============");
                    }

                }
            }
        
            //  string cabecalho = listaevents.Select(w => w.Id == nrale2);
          
            
            Console.ReadKey();
        }
    }
}
