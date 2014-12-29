BEGIN {
	FS = ";"
	OFS = ";"
}

{				
	sub(/\,/, "\." , $4)	# msrp
	sub(/\,/, "\." , $5)	# refp 
	sub(/\,/, "\." , $6)	# xferbasep
	sub(/\,/, "\." , $8)	# lmsrp_ru
	print
}
