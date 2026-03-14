/*
List of distinct rsidR values with count for each.

You can then use a particular rsidR value in another query to list all the files with that value.
Also list their rsidRoot value in that query to see evidence of a common rsidR value common
across different documents.

*/

SELECT DISTINCT rsid_value, count(rsid_value) AS tot_rsids
FROM rsids
WHERE rsid_type == "rsidR"
GROUP BY rsid_value
ORDER BY tot_rsids DESC