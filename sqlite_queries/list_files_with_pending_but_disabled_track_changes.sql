SELECT DISTINCT ds.file_name, COUNT(tc.file_name) AS "pending revisions"
FROM document_summary AS ds
JOIN track_changes AS tc ON ds.file_name == tc.file_name
WHERE ds.track_changes_enabled = 0
GROUP BY ds.file_name
ORDER BY "pending revisions" DESC
