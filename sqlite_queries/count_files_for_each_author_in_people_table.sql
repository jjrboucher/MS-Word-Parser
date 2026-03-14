/*
List of all authors in the people table, with a count of how many times they appear.
*/

SELECT author, count(DISTINCT file_name) AS "files with this author"
FROM people
GROUP BY author
ORDER BY "files with this author" DESC