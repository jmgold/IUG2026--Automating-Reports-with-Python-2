/* Weekly New Report
The report will retrieve all items, not in the specifically excluded statuses
('m','n','z','t','s','$','d','8','w','y'), that were created in the last 10 days.
It will also include a count of items per bib, count of orders with a status of "o",
and a count of bib-level holds.
 */

SELECT 
  distinct 'b'|| rmb.record_num || 'a' AS "Bib Record Num",
  i.location_code, 
  CASE
    WHEN pei.index_entry IS NULL THEN UPPER(peb.index_entry) 
    ELSE UPPER(pei.index_entry)
  END AS "Call#",
  brp.best_author AS "Author",
  brp.best_title AS "Title",
  string_agg(distinct i.barcode, ' ') AS "Barcode", 
  string_agg(distinct pes.index_entry, ' | ') AS "Series Info",
  count(distinct ic.id) AS "Item Count",  
  count(distinct o.id) AS "Order Count",
  count(distinct h.id) AS "Hold Count"
FROM sierra_view.item_view i
JOIN sierra_view.bib_record_item_record_link bri
  ON i.id = bri.item_record_id
JOIN sierra_view.record_metadata rmb
  ON bri.bib_record_id = rmb.id
  AND rmb.record_type_code='b'
JOIN sierra_view.phrase_entry peb
  ON i.id = peb.record_id
  AND peb.index_tag='c'
JOIN sierra_view.bib_record_property brp
  ON brp.bib_record_id = bri.bib_record_id
LEFT JOIN sierra_view.phrase_entry pei
  ON pei.record_id=bri.item_record_id
  AND pei.index_tag='c'
LEFT JOIN sierra_view.phrase_entry pes
  ON pes.record_id=bri.bib_record_id
  AND pes.index_tag='t'
  AND pes.varfield_type_code='s'
LEFT JOIN sierra_view.item_record ic
  ON ic.id=i.id
  AND ic.item_status_code not in ('m','n','z','t','s','$','d','8','w','y')
LEFT JOIN sierra_view.hold h
  ON (bri.bib_record_id = h.record_id OR i.id = h.record_id)
LEFT JOIN sierra_view.bib_record_order_record_link bro
  ON bri.bib_record_id = bro.bib_record_id
LEFT JOIN sierra_view.order_record o
  ON bro.order_record_id = o.id
  AND o.order_status_code ='o'

WHERE i.record_creation_date_gmt::date>=NOW()::DATE-EXTRACT(DOW FROM NOW())::INTEGER-10 
GROUP BY rmb.record_num, i.location_code, "Call#", brp.best_author, brp.best_title
ORDER BY i.location_code, "Call#", brp.best_author, brp.best_title
