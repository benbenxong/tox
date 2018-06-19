select object_name oname,object_id,created,object_type from all_objects where rownum<50000;

select object_name oname,created,object_type from all_objects where rownum<50
;
