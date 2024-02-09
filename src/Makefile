
.PHONY: all
all: addin

.PHONY: addin
addin: bin/PresentationCloseFinalAddin.ppam

MODULE = Module.vb
CLASS = AppClass.vb

bin/PresentationCloseFinalAddin.ppam: $(MODULE) $(CLASS)
	vbamc -m $(MODULE) -c $(CLASS) -n "PresentationCloseFinalAddin" --company "Acme" -f PresentationCloseFinalAddin.ppam

.PHONY: clean
clean:
	rm -rf bin/ obj/
