/*
List of files with a revision count of 1, meaning it's an original document.
*/
SELECT file_name, revision
FROM metadata
WHERE revision == 1