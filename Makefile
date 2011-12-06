all: clean rack_layout

rack_layout:
	perl ./bin/rack_layout.pl

clean:
	rm -rf reports/*.xls
