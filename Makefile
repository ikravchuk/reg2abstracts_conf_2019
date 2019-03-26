## run	: run abstract.py 
.PHONY : run
run : abstract.py
	@python abstract.py


## help	:	display this help message
.PHONY : help
help : Makefile
	@sed -n 's/^##//p' $<
