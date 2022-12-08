# paper_reference_entry
This python script checks for a given citation format from a citation list sourced from a column of xlsx extension, if found then it gives the earlier used reference number, if not found then enters the new entry as the latest reference. Ofcourse online tools like overleaf addresses the duplicacy issue of reference like a charm but whoever is still writing paper on a doc, its for them :p

python package used: openpyxl 
	
xlsx format-->

columns: RefSL, p/w (paper or website), citation (harvard format from google scholar), year, DOI/link

rows: each rows indicate separate reference citation
