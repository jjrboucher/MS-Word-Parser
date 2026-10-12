SELECT author, COUNT(*) AS "total count"
FROM track_changes
WHERE author IS NOT NULL
GROUP BY author
ORDER BY "total count" DESC
