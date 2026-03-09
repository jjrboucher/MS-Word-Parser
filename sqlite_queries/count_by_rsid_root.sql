/*
Simple query to get a count of how many files there are for a given rsidRoot value
*/
SELECT rsid_root, COUNT(*) AS "total count"
FROM document_summary
GROUP BY rsid_root
ORDER BY "total count" DESC
