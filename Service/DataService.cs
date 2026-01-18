using System;
using ClosedXML.Excel;    
using Microsoft.Toolkit.Uwp.Notifications;
using System.IO;
using System.Collections.Generic; // Manipulação de listas e coleções
using System.Linq;                // Consultas e operações em coleções (Skip, Where, etc)
using System.Xml;                 // Manipulação básica de XML
using System.ComponentModel;
using Model.Users;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
namespace Service.DataServices{
public class DataService
    {
        public static List<RegistroLogistica> UsuariosCarregados  = new();
        public static List<RegistroLogistica> NomesVazio = new();
        public static bool valorcondicional;
            private static string? UserPath;
        private static string? CaminhoFinal;

        public DataService(){
            UserPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        CaminhoFinal = Path.Combine(UserPath, @"OneDrive - CMPC\Descarga Materiais.xlsx");
        }
     public async Task GetDataSheet()
        {
            await Task.Run(()=>{
                UsuariosCarregados.Clear();
            using(var stream = new FileStream(CaminhoFinal, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                
                using (var workbook = new XLWorkbook(stream))
        {
            var tabelas=workbook.Worksheet("Tabela2");
            var linhas = tabelas.RowsUsed();

            List<int> cells=new List<int>{3,4,5,7,8,10,6,21,9,11};
            

         foreach (var linha in linhas.Skip(1))
{
    var registro = new RegistroLogistica();
    registro.numLinhas=linha.RowNumber();

    // Percorrendo sua lista de índices
    foreach (var colIndex in cells)
    {
        var valor = linha.Cell(colIndex).GetValue<string>();

        try{
        switch (colIndex)
        {
            case 3:  registro.Cavalo = valor; break;
            case 4:  registro.Carreta = valor; break;
            case 5:  registro.Carreta2 = valor; break;
            case 6:  registro.NomeCompleto = valor; break;
            case 7:  registro.Telefone = valor; break;
            case 8:   
            registro.Transportadora = valor; break;
            case 9: registro.Fornecedor = valor;break; 
            case 10:registro.NotaFiscal = valor; break;
            
            case 11: registro.Horario = valor; break;
            case 21: registro.RG = valor; break; 
        }
        }
         catch
            {
                    new ToastContentBuilder()
                    .AddText($"Erro ao registrar dados da planilha")
                    .Show();               
            }

    }

    UsuariosCarregados.Add(registro);

    var ultimo= UsuariosCarregados.LastOrDefault();
    var LinhasVazias=linhas.Where(l=>string.IsNullOrEmpty(l.Cell(21).GetValue<string>()));
      foreach(var u in UsuariosCarregados){
      if(valorcondicional &&  string.IsNullOrEmpty(u.NomeCompleto)) {
           new ToastContentBuilder()
      .AddText($"Nomes dos Motoristas:{u.NomeCompleto} ")
      .AddText($"Placa do Motorista:{u.Cavalo}").Show();
        }

        
      }
}
}      
        
            }
    
    
    
            });
        }

    public async Task FiltraUser()
        {
            await Task.Run(()=>{
            var rgigual= UsuariosCarregados.GroupBy(u=>u.RG).Where(u=>u.Count()>1).ToList();
           foreach(var grupo in rgigual){
            var CadastroCompleto=grupo.FirstOrDefault(u => !string.IsNullOrEmpty(u.NomeCompleto));
            var cadastrosVazios = grupo.Where(u => string.IsNullOrEmpty(u.NomeCompleto));

            if (CadastroCompleto != null)
            {
               foreach(var vazio in cadastrosVazios){
          var preencher= new RegistroLogistica();
          preencher.numLinhas = vazio.numLinhas;
    preencher.RG = vazio.RG;
    preencher.NotaFiscal = vazio.NotaFiscal;
    preencher.Horario = vazio.Horario; 
    preencher.NomeCompleto = CadastroCompleto.NomeCompleto;
    preencher.Telefone = CadastroCompleto.Telefone;
    preencher.Cavalo = CadastroCompleto.Cavalo;    
    preencher.Carreta=CadastroCompleto.Carreta;  
    preencher.Carreta2 = CadastroCompleto.Carreta2;
    preencher.Categoria = CadastroCompleto.Categoria;
    preencher.Transportadora = CadastroCompleto.Transportadora;
    preencher.Fornecedor = CadastroCompleto.Fornecedor;

    NomesVazio.Add(preencher);
                
               }
            }
           }   
            });

            }
            public async Task SalvarPreenchimento()
        {
            await Task.Run(()=>{
         using (var workbook = new XLWorkbook(CaminhoFinal))
    {
        var planilha = workbook.Worksheet("Tabela2");
        
        // Localizamos a última linha usada para começar a escrever na de baixo

        foreach (var registro in NomesVazio)
        {
        int numLinha = registro.numLinhas??0;

            // Aqui você deve seguir a ordem exata das colunas da sua planilha original
            // Exemplo baseado no seu switch case:
            planilha.Cell(numLinha, 3).Value = registro.Cavalo;
            planilha.Cell(numLinha, 4).Value = registro.Carreta;
            planilha.Cell(numLinha, 5).Value = registro.Carreta2;
            planilha.Cell(numLinha, 6).Value = registro.NomeCompleto;
            planilha.Cell(numLinha, 7).Value = registro.Telefone;
            planilha.Cell(numLinha, 8).Value = registro.Transportadora; // E Transportadora se for a mesma célula
            planilha.Cell(numLinha, 9).Value = registro.Fornecedor;
            planilha.Cell(numLinha, 21).Value = registro.RG;

            // Muda a cor da linha ou coloca um aviso para você saber que foi preenchido via código
            planilha.Row(numLinha).Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF2F8"); 

        }

        // Salva as alterações no mesmo arquivo referenciado
        workbook.Save();
        System.IO.File.SetLastWriteTime(CaminhoFinal,DateTime.Now);
         var tabelas=workbook.Worksheet("Tabela2");
            var linhas = tabelas.RowsUsed();
             var ultimo= UsuariosCarregados.LastOrDefault();
    var LinhasVazias=linhas.Where(l=>string.IsNullOrEmpty(l.Cell(21).GetValue<string>()));

        if(valorcondicional && LinhasVazias!=null && ultimo !=null) new ToastContentBuilder()
        .AddText($"\n[SUCESSO]").Show();

        
    }
        });
        }
             public async Task OpcaodeMensagem()
        {

            using(var workbook = new XLWorkbook(CaminhoFinal))
            {
                var planilha= workbook.Worksheet("Tabela2");
                var cell= planilha.Cell("V18");
                 valorcondicional=!cell.IsEmpty() && cell.Value.IsBoolean && cell.GetBoolean();

            }
        }
        }
    }
