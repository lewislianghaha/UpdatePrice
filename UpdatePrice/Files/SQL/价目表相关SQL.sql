SELECT *
from dbo.ICPrcPly


SELECT A.FInterID,A.FName
from dbo.ICPrcPly A

SELECT b.FCheckerID,b.FCheckDate,b.FChecked,b.FModel,b.*
from dbo.ICPrcPlyEntry b
WHERE FItemID in(127255,127337,127653)
AND FInterID=2
AND b.FRelatedID=91428


SELECT TOP 1 A.FInterID,A.FRelatedID,A.FItemID,A.FBegDate,A.FEndDate,A.FPrice,A.FCheckerID,A.FChecked,A.FCheckDate,A.FMainterID,A.FMaintDate
FROM dbo.ICPrcPlyEntry A


UPDATE ICprc SET Icprc.FBegDate=@FBegDate,Icprc.FEndDate=@FEndDate,Icprc.FPrice=@FPrice,Icprc.FCheckerID=16394,Icprc.FChecked=1,Icprc.FCheckDate=@FCheckDate,Icprc.FMainterID=16394,Icprc.FMaintDate=@FCheckDate
FROM ICPrcPlyEntry ICprc
WHERE ICprc.FInterID=@FInterID
AND ICprc.FRelatedID=@FCustID
AND ICprc.FItemID=@FItemID





SELECT DISTINCT B.FItemID,B.FName
from dbo.ICPrcPlyEntry A
INNER JOIN dbo.t_Organization B ON A.FRelatedID=B.FItemID
WHERE A.FInterID=2


SELECT *
from dbo.t_ICItem 
WHERE FItemID =12725566

SELECT *
from dbo.t_Organization

SELECT A.FInterID,A.FName,
	   b.FInterID,b.FItemID,b.FRelatedID,b.FPrice,b.FBegDate,b.FEndDate,
	   b.FChecked,b.FCheckerID,b.FCheckDate,
	   b.FMaintDate,b.FMainterID
from dbo.ICPrcPly a
INNER JOIN dbo.ICPrcPlyEntry b ON a.FInterID=b.FInterID
WHERE a.FInterID=2



SELECT DISTINCT B.FRelatedID,C.FName
FROM dbo.ICPrcPly A
INNER JOIN dbo.ICPrcPlyEntry B ON A.FInterID=B.FInterID
INNER JOIN dbo.t_Organization C ON B.FRelatedID=C.FItemID
WHERE A.FInterID=2
