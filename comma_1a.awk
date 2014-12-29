BEGIN {
	FS = ";"
	OFS = ";"
}

{				
	sub(/\,/, "" , $4)	 # msrp - удаление запятой в разрядах
	sub(/\./, "\," , $4) # msrp - замена десятичной точки на запятую
	sub(/\,/, "" , $5)	 # refp
	sub(/\./, "\," , $5) # refp
	sub(/\,/, "" , $6)	 # xferbasep	
	sub(/\./, "\," , $6) # xferbasep
	print
}
