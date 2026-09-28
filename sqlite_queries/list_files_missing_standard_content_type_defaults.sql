SELECT DISTINCT file_name
FROM content_types
WHERE file_name NOT IN (
    SELECT file_name FROM content_types WHERE type = "Default" AND extension_part_name = "rels"
)
OR file_name NOT IN (
    SELECT file_name FROM content_types WHERE type = "Default" AND extension_part_name = "xml"
)
