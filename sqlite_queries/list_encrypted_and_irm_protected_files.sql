SELECT ds.file_name, ds.encrypted, rm.owner, rm.organization, rm.template_name
FROM document_summary AS ds
LEFT JOIN rights_management AS rm ON ds.file_name == rm.file_name
WHERE ds.encrypted = 1
ORDER BY ds.file_name
