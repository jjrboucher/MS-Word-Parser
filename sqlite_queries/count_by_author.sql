/*
Simple query to get a count of how many files there are for a given author.
*/
SELECT author, COUNT(*) AS "total count"
FROM metadata
GROUP BY author
HAVING COUNT(*) > 1
ORDER BY "total count" DESC