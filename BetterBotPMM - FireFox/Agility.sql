SELECT REPLACE(SUM(case when Agility.Tipo = 'IMP' and year(CURRENT_TIMESTAMP) = year(Agility.Data) then Agility.Desconsolidacao else 0 end),'.',',') as AgilityAno,
REPLACE(SUM(case when Agility.Tipo = 'IMP' and year(CURRENT_TIMESTAMP) = year(Agility.Data) and MONTH(current_timestamp) = month(Agility.Data) then Agility.Desconsolidacao else 0 end),'.',',') as AgilityMes
from Agility