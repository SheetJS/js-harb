SHELL=/bin/bash
LIB=harb
FMT=csv dif sylk prn socialcalc dbf misc full
REQS=
ADDONS=dist/cpharb.js
AUXTARGETS=
CMDS=bin/harb.njs
HTMLLINT=

ULIB=$(shell echo $(LIB) | tr a-z A-Z)
DEPS=$(sort $(wildcard bits/*.js))
TARGET=$(LIB).js
FLOWTARGET=$(LIB).flow.js
FLOWTGTS=$(TARGET) $(AUXTARGETS)
UGLIFYOPTS=--support-ie8

## Main Targets

.PHONY: all
all: $(TARGET) $(AUXTARGETS) ## Build library and auxiliary scripts

$(FLOWTGTS): %.js : %.flow.js
	node -e 'process.stdout.write(require("fs").readFileSync("$<","utf8").replace(/^[ \t]*\/\*[:#][^*]*\*\/\s*(\n)?/gm,"").replace(/\/\*[:#][^*]*\*\//gm,""))' > $@

$(FLOWTARGET): $(DEPS)
	cat $^ | tr -d '\15\32' > $@

bits/01_version.js: package.json
	echo "$(ULIB).version = '"`grep version package.json | awk '{gsub(/[^0-9a-z\.-]/,"",$$2); print $$2}'`"';" > $@


.PHONY: clean
clean: clean-baseline ## Remove targets and build artifacts
	rm -f $(TARGET) $(FLOWTARGET)

.PHONY: init
init: ## Initial setup for development
	git submodule init
	git submodule update
	git submodule foreach git pull origin master
	git submodule foreach make
	mkdir -p tmp

.PHONY: dist
dist: dist-deps $(TARGET) ## Prepare JS files for distribution
	cp $(TARGET) dist/
	cp LICENSE dist/
	uglifyjs $(UGLIFYOPTS) $(TARGET) -o dist/$(LIB).min.js --source-map dist/$(LIB).min.map --preamble "$$(head -n 1 bits/00_header.js)"
	misc/strip_sourcemap.sh dist/$(LIB).min.js
	uglifyjs $(UGLIFYOPTS) $(REQS) $(TARGET) -o dist/$(LIB).core.min.js --source-map dist/$(LIB).core.min.map --preamble "$$(head -n 1 bits/00_header.js)"
	misc/strip_sourcemap.sh dist/$(LIB).core.min.js
	uglifyjs $(UGLIFYOPTS) $(REQS) $(ADDONS) $(TARGET) -o dist/$(LIB).full.min.js --source-map dist/$(LIB).full.min.map --preamble "$$(head -n 1 bits/00_header.js)"
	misc/strip_sourcemap.sh dist/$(LIB).full.min.js

.PHONY: dist-deps
dist-deps:
	cp misc/cpharb.js dist/cpharb.js

.PHONY: codepage
codepage: ## Rebuild cpharb.js codepage script
	./misc/make_cp.sh


## Testing

.PHONY: test mocha
test mocha: test.js $(TARGET) baseline ## Run test suite
	mocha -R spec -t 30000

#*                      To run tests for one format, make test_<fmt>
TESTFMT=$(patsubst %,test_%,$(FMT))
.PHONY: $(TESTFMT)
$(TESTFMT): test_%:
	FMTS=$* make test

.PHONY: baseline
baseline: ## Build test baselines
	#

.PHONY: clean-baseline
clean-baseline: ## Remove test baselines
	#rm -f *.csv

## Code Checking

.PHONY: lint
lint: $(TARGET) $(AUXTARGETS) ## Run jshint and jscs checks
	@jshint --show-non-errors $(TARGET) $(AUXTARGETS)
	@jshint --show-non-errors $(CMDS)
	@jshint --show-non-errors package.json
	@jshint --show-non-errors --extract=always $(HTMLLINT)
	@jscs $(TARGET) $(AUXTARGETS)

.PHONY: flow
flow: lint ## Run flow checker
	@flow check --all --show-all-errors

.PHONY: cov
cov: misc/coverage.html ## Run coverage test

#*                      To run coverage tests for one format, make cov_<fmt>
COVFMT=$(patsubst %,cov_%,$(FMT))
.PHONY: $(COVFMT)
$(COVFMT): cov_%:
	FMTS=$* make cov

misc/coverage.html: $(TARGET) test.js
	mocha --require blanket -R html-cov -t 20000 > $@

.PHONY: coveralls
coveralls: ## Coverage Test + Send to coveralls.io
	mocha --require blanket --reporter mocha-lcov-reporter -t 20000 | node ./node_modules/coveralls/bin/coveralls.js


.PHONY: help
help:
	@grep -hE '(^[a-zA-Z_-][ a-zA-Z_-]*:.*?|^#[#*])' $(MAKEFILE_LIST) | bash misc/help.sh

#* To show a spinner, append "-spin" to any target e.g. cov-spin
%-spin:
	@make $* & bash misc/spin.sh $$!
