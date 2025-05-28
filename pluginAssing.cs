using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

public class AtribuirEquipePorPlataforma : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        var service = serviceFactory.CreateOrganizationService(context.UserId);
        var tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

        tracing.Trace("Plugin AtribuirEquipePorPlataforma iniciado.");

        if (!context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is Entity target))
        {
            tracing.Trace("Target não encontrado ou inválido.");
            return;
        }

        if (target.LogicalName != "vc4_registrodevisita")
        {
            tracing.Trace($"Entidade não suportada: {target.LogicalName}");
            return;
        }

        try
        {
            var plataformaOption = GetPlataformaOption(target, tracing);
            if (plataformaOption == null)
            {
                tracing.Trace("Campo vc4_plataforma não preenchido. Encerrando plugin.");
                return;
            }

            var equipesMapeadas = ObterEquipesDasPlataformas(service, tracing);
            if (equipesMapeadas == null || !equipesMapeadas.Any())
            {
                tracing.Trace("Nenhum valor encontrado na entidade VC4_Parametros para o parâmetro 'Equipes das Plataformas'.");
                return;
            }

            var equipeId = BuscarEquipePorPlataforma(equipesMapeadas, plataformaOption.Value, tracing);
            if (string.IsNullOrWhiteSpace(equipeId))
            {
                tracing.Trace($"Não foi encontrado um mapeamento de equipe para a plataforma: {plataformaOption.Value}");
                return;
            }

            AtribuirRegistroParaEquipe(target, equipeId, tracing);
            tracing.Trace($"Registro atribuído com sucesso para a equipe {equipeId}.");

        }
        catch (Exception ex)
        {
            tracing.Trace($"Erro no plugin AtribuirEquipePorPlataforma: {ex}");
            // Não lança exceção — Fail Safe
        }

        tracing.Trace("Plugin AtribuirEquipePorPlataforma finalizado.");
    }

private OptionSetValue GetPlataformaOption(Entity target, IOrganizationService service, ITracingService tracing)
{
    try
    {
        if (!target.Attributes.Contains("vc4_revenda") || !(target["vc4_revenda"] is EntityReference lojistaRef))
        {
            tracing.Trace("O campo vc4_revenda não está preenchido ou não é um EntityReference.");
            return null;
        }

        tracing.Trace($"Buscando plataforma do lojista com ID: {lojistaRef.Id}");

        var lojista = service.Retrieve("vc4_lojistas", lojistaRef.Id, new ColumnSet("vc4_plataforma"));

        if (lojista == null)
        {
            tracing.Trace($"Nenhum lojista encontrado com ID {lojistaRef.Id}");
            return null;
        }

        if (!lojista.Attributes.Contains("vc4_plataforma"))
        {
            tracing.Trace($"O lojista {lojistaRef.Id} não possui o campo vc4_plataforma preenchido.");
            return null;
        }

        var plataformaOption = lojista.GetAttributeValue<OptionSetValue>("vc4_plataforma");

        if (plataformaOption != null)
            tracing.Trace($"Plataforma encontrada no lojista: {plataformaOption.Value}");
        else
            tracing.Trace("Plataforma do lojista é nula.");

        return plataformaOption;
    }
    catch (Exception ex)
    {
        tracing.Trace($"Erro ao obter plataforma do lojista: {ex.Message}");
        return null;
    }
}


    private List<Dictionary<string, string>> ObterEquipesDasPlataformas(IOrganizationService service, ITracingService tracing)
    {
        try
        {
            var query = new QueryExpression("vc4_parametros")
            {
                ColumnSet = new ColumnSet("vc4_conteudo"),
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("vc4_name", ConditionOperator.Equal, "Equipes das Plataformas")
                    }
                }
            };

            var result = service.RetrieveMultiple(query);
            if (result.Entities.Count == 0)
            {
                tracing.Trace("Parâmetro 'Equipes das Plataformas' não encontrado na tabela VC4_Parametros.");
                return null;
            }

            var json = result.Entities[0].GetAttributeValue<string>("vc4_conteudo");
            if (string.IsNullOrWhiteSpace(json))
            {
                tracing.Trace("O campo vc4_conteudo está vazio.");
                return null;
            }

            var equipes = JsonSerializer.Deserialize<List<Dictionary<string, string>>>(json);
            return equipes;
        }
        catch (Exception ex)
        {
            tracing.Trace($"Erro ao obter parâmetros das plataformas: {ex.Message}");
            return null;
        }
    }

    private string BuscarEquipePorPlataforma(List<Dictionary<string, string>> equipes, int plataformaValue, ITracingService tracing)
    {
        try
        {
            var key = plataformaValue.ToString();

            var equipe = equipes.FirstOrDefault(e => e.ContainsKey(key));
            if (equipe != null)
            {
                return equipe[key];
            }

            return null;
        }
        catch (Exception ex)
        {
            tracing.Trace($"Erro ao buscar equipe para plataforma {plataformaValue}: {ex.Message}");
            return null;
        }
    }

    private void AtribuirRegistroParaEquipe(Entity target, string teamId, ITracingService tracing)
    {
        try
        {
            var equipeGuid = Guid.Parse(teamId);
            target["ownerid"] = new EntityReference("team", equipeGuid);
            tracing.Trace($"OwnerId atribuído para equipe {equipeGuid}.");
        }
        catch (Exception ex)
        {
            tracing.Trace($"Erro ao atribuir OwnerId: {ex.Message}");
        }
    }
}
