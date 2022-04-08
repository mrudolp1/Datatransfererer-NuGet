SELECT
    sm.foundation_group_id
FROM
    gen.model_work_order_xref wox
    ,gen.structure_model sm
WHERE
    wox.work_order_seq_num=[WO]
    AND wox.model_id=sm.ID

