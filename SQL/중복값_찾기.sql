SELECT "SR_No_ATTR", COUNT(*)
FROM "표준데이터시트_개별속성_240528"
GROUP BY "SR_No_ATTR"
HAVING COUNT(*) > 1 ;
