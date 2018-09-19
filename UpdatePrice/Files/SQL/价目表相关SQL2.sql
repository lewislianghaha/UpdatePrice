SELECT /*b.FCheckerID,b.FCheckDate,b.FChecked,*/b.*
from dbo.ICPrcPlyEntry b
WHERE /*FItemID =(*//*127255,127337,127653)
AND */FInterID=1
AND b.FRelatedID=0
AND B.FItemID=127255


SELECT *
from dbo.t_ICItem 
WHERE FItemID =127259


SELECT A.FInterID,A.FRelatedID,A.FItemID,A.FBegDate,A.FEndDate,A.FPrice,A.FCheckerID,A.FChecked,A.FCheckDate,A.FMainterID,A.FMaintDate
from dbo.ICPrcPlyEntry a
WHERE a.FItemID in(127255,127337,127653)  --127255
AND a.FInterID=2
AND a.FRelatedID=91428

SELECT a.FName
from dbo.t_Organization a
WHERE a.FItemID=91428


