using System;
using System.Threading.Tasks; // Necessário para Task
using Model.Users;
using Service.DataServices;

namespace Programa
{
    class Program
    {
        // Transformamos o Main em async para poder usar await
        public static async Task Main(string[] args)
        {
            

            // 1. Instanciamos o serviço
            var servico = new DataService();
            // 2. Chamamos o método e ESPERAMOS ele terminar (await)
            // Se não usar await, o programa pula para o próximo comando antes de ler o arquivo
            await servico.GetDataSheet();
            await servico.FiltraUser();
            await servico.OpcaodeMensagem();
            await servico.SalvarPreenchimento();


            // 3. Agora sim, verificamos a lista que foi preenchida na DataService
          

            // 4. Percorrendo os dados
            
            
            

            await servico.GetDataSheet();


/*foreach(var reg in DataService.NomesVazio) {
    Console.WriteLine($"RG: {reg.RG} - Nome Recuperado: {reg.NomeCompleto}");
}*/


        }
   
    }
}