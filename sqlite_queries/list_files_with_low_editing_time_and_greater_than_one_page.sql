/*
List files with a very low editing time (1 minute or less) but more than one page long.
This could help establish whether the content was copied/pasted into the document, or if
generated automatically via an online tool (AI, document creating service).
*/

SELECT file_name, pages, total_editing_time
FROM metadata
WHERE total_editing_time IN ("0", "1") AND pages >1