WITH
  -- First Table: Data from ps_mstr (5 columns)
  PS_Data AS (
    SELECT 
      '2674' AS [Plant],
      [ps_par],  
      [ps_comp],  
      [ps_qty_per],  
      [ps_rmks]
    FROM 
      [QADEE].[dbo].[ps_mstr]
    WHERE 
      [ps_end] IS NULL
    UNION ALL
    SELECT 
      '2798' AS [Plant],
      [ps_par],  
      [ps_comp],  
      [ps_qty_per],  
      [ps_rmks]
    FROM 
      [QADEE2798].[dbo].[ps_mstr]
    WHERE 
      [ps_end] IS NULL
  ),

  -- Second Table: Data from PodData + pt_mstr (25 columns)
  PodData AS (
    SELECT  
      pd.[pod_po_site],  
      pd.[pod__chr08],  
      pm.[po_vend],  
      pd.[pod_nbr],  
      pd.[pod_line],  
      pd.[pod_part],  
      pd.[pod_cum_qty[1]]],  
      pd.[pod_ord_mult],  
      pd.[pod_translt_days],
      pd.[pod_start_eff[1]]],
      pd.[pod_curr_rlse_id[1]]],
      pt.[pt_desc1],
      pt.[pt_desc2],
      pt.[pt_prod_line],
      pt.[pt_group],
      pt.[pt_part_type],
      pt.[pt_status],
      pt.[pt_abc],
      pt.[pt_cyc_int],
      pt.[pt_sfty_stk],
      pt.[pt_sfty_time],
      pt.[pt_buyer],
      pt.[pt_vend],
      pt.[pt__chr02],
      pt.[pt_dsgn_grp]
    FROM  
      [QADEE].[dbo].[pod_det] pd  
      JOIN [QADEE].[dbo].[po_mstr] pm ON pd.[pod_nbr] = pm.[po_nbr]  
      LEFT JOIN [QADEE].[dbo].[pt_mstr] pt 
        ON pd.[pod_po_site] = pt.[pt_site] 
        AND pd.[pod_part] = pt.[pt_part]
    WHERE  
      pd.[pod_end_eff[1]]] = '2049-12-31 00:00:00'  
    UNION ALL  
    SELECT  
      pd.[pod_po_site],  
      pd.[pod__chr08],  
      pm.[po_vend],  
      pd.[pod_nbr],  
      pd.[pod_line],  
      pd.[pod_part],  
      pd.[pod_cum_qty[1]]],  
      pd.[pod_ord_mult],  
      pd.[pod_translt_days],
      pd.[pod_start_eff[1]]],
      pd.[pod_curr_rlse_id[1]]],
      pt.[pt_desc1],
      pt.[pt_desc2],
      pt.[pt_prod_line],
      pt.[pt_group],
      pt.[pt_part_type],
      pt.[pt_status],
      pt.[pt_abc],
      pt.[pt_cyc_int],
      pt.[pt_sfty_stk],
      pt.[pt_sfty_time],
      pt.[pt_buyer],
      pt.[pt_vend],
      pt.[pt__chr02],
      pt.[pt_dsgn_grp]
    FROM  
      [QADEE2798].[dbo].[pod_det] pd  
      JOIN [QADEE2798].[dbo].[po_mstr] pm ON pd.[pod_nbr] = pm.[po_nbr]  
      LEFT JOIN [QADEE2798].[dbo].[pt_mstr] pt 
        ON pd.[pod_po_site] = pt.[pt_site] 
        AND pd.[pod_part] = pt.[pt_part]
    WHERE  
      pd.[pod_end_eff[1]]] = '2049-12-31 00:00:00'
  )

-- Final Merged Result (30 Columns)
SELECT
  -- First 5 columns from PS_Data
  ps.[Plant],
  ps.[ps_par],
  ps.[ps_comp],
    pd.[pt_desc1],
  pd.[pt_desc2],
  ps.[ps_qty_per],
  
  -- Next 25 columns from PodData
  pd.[pod__chr08],
  pd.[po_vend],
  pd.[pod_nbr],
  pd.[pod_line],
  pd.[pod_cum_qty[1]]],
  pd.[pod_ord_mult],
  pd.[pod_translt_days],
  pd.[pod_start_eff[1]]],
  pd.[pod_curr_rlse_id[1]]],

  pd.[pt_prod_line],
  pd.[pt_group],
  pd.[pt_status],
  pd.[pt_sfty_stk],
  pd.[pt_sfty_time],
  pd.[pt_buyer],
  pd.[pt_vend],
  pd.[pt__chr02],
  pd.[pt_dsgn_grp]
FROM
  PS_Data ps
  LEFT JOIN PodData pd 
    ON ps.[Plant] = pd.[pod_po_site] 
    AND ps.[ps_comp] = pd.[pod_part];