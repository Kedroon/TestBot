SELECT SUM(t.Desconsolidacao) AS total_Valor
    FROM (SELECT Desconsolidacao FROM Capital where Tipo = 'IMP'
          UNION ALL
          SELECT Desconsolidacao FROM Agility where Tipo = 'IMP'
		  UNION ALL
          SELECT Desconsolidacao FROM UPS where Tipo = 'IMP'
		  UNION ALL
          SELECT Desconsolidacao FROM KN where Tipo = 'IMP'
		  UNION ALL
          SELECT Desconsolidacao FROM EXPEDITORS where Tipo = 'IMP'
		  UNION ALL
		  SELECT Desconsolidacao FROM Nippon where Tipo = 'IMP') t