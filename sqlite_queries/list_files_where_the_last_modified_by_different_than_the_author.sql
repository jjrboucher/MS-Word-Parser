/*
List of files where the author and last modified by is different.
*/

SELECT file_name, author, last_modified_by
FROM metadata
WHERE author != last_modified_by