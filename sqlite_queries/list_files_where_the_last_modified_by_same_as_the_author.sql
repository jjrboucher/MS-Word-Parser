/*
List of files where the author and last modified by is the same.
*/

SELECT file_name
FROM metadata
WHERE author == last_modified_by