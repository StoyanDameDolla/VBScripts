str = "Stoyan Steven Sylvester Sam Samir"

S[a-z]{1,} 'similar symbol is +
S[a-z]{0,} 'similar symbol is *

str_1 = "12311 123124 53657 23368756"  

[0-9]{5} [0-9]{6} [0-9]{5} [0-9]{8}			'alternative method --> \d{5}\d{6}\d{5}\d{8}

str_2 = "Ann is cool person"

[\bAnn\b]

str_3 = "Ann is cool person"

\b
'^ --> represents the beginning of a string
'^$ --> capture an entire stream of characters
'^(.|\n)*$
'\?--> checks exactly for the question mark sign

