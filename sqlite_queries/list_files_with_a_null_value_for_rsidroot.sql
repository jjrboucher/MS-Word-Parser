/*
See all files where the rsidRoot value is NULL
*/

SELECT file_name
FROM document_summary
WHERE rsid_root ISNULL
ORDER BY file_name