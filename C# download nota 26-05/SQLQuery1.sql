select SUM(case when Periodo <> 1 and Tipo = 'Importação' then Valor else null end) as ImportacaoExtra , 
SUM(case when Periodo <> 1 and Tipo = 'Exportação' then Valor else null end) as ExportacaoExtra , 
SUM(case when Tipo = 'Importação' then Valor else null end) as TotalImportacao, 
SUM(case when Tipo = 'Exportação' then Valor else null end) as TotalExportacao
from Notas