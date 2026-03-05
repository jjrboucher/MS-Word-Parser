/*
Simple query to get a count of how many files there are for a given author.
*/
SELECT last_modified_by, COUNT(*) AS "total count"
FROM metadata
GROUP BY last_modified_by
HAVING COUNT(*) > 1
ORDER BY "total count" DESC

SELECT author, COUNT(*) AS "total count"
FROM metadata
GROUP BY author
HAVING COUNT(*) > 1
ORDER BY "total count" DESC