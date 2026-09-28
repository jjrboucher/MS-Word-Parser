SELECT editor, COUNT(DISTINCT file_name) AS "files", COUNT(*) AS "ranges"
FROM range_permissions
WHERE editor IS NOT NULL
GROUP BY editor
ORDER BY "files" DESC
