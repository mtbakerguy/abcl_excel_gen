FILE=xyz.xlsx

JAVA=java
ABCL=$(HOME)/abcl_excel_gen/abcl-bin-1.6.0/abcl.jar
POI=$(HOME)/abcl_excel_gen/poi-4.1.1/poi-4.1.1.jar

%.xlsx : %.lisp buildsheet.lisp
	$(JAVA) -jar $(ABCL) --noinform -- $(POI) $@ $< < buildsheet.lisp

all: $(FILE)

clean:
	rm -fr $(FILE)

