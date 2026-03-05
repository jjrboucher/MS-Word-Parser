/*
Search for files that were created or edited by certain individuals
*/
SELECT file_name, author, last_modified_by
FROM metadata
WHERE author IN ("<name_1>", "<name_2>") OR last_modified_by IN ("<name_1>", "<name_2>")

SELECT file_name, author, last_modified_by
FROM metadata
WHERE (author LIKE '%<name_1>%' OR author LIKE '%<name_2>%')
   OR (last_modified_by LIKE '%<name_1>%' OR last_modified_by LIKE '%<name_2>%')