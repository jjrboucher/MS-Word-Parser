/*
Simple query to get a list of applications and count
*/
SELECT application, COUNT(*) AS "total count"
FROM metadata
GROUP BY application
ORDER BY "total count" DESC