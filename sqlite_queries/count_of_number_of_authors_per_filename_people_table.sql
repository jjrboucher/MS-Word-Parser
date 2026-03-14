/*
This query returns a list of files and count of # of author for each file as found in the people table.
*/

SELECT file_name, COUNT(DISTINCT author) AS "author count"
FROM people
GROUP BY file_name
ORDER BY "author count" DESC