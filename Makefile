.PHONY: test test-roundtrip roundtrip-example clean-roundtrip-example

MANUAL_DIR := artifacts
TEST_IN := Preregistration_Original.docx
TEST_MD := $(MANUAL_DIR)/out_test.md
TEST_DOCX := $(MANUAL_DIR)/out_test.docx
ABS_TEST_IN := $(abspath $(TEST_IN))
ABS_TEST_MD := $(abspath $(TEST_MD))
ABS_TEST_DOCX := $(abspath $(TEST_DOCX))

test: roundtrip-example
	python3 -m unittest -q

test-roundtrip:
	python3 -m unittest -q tests.test_roundtrip_example tests.test_roundtrip_edges

roundtrip-example:
	mkdir -p $(MANUAL_DIR)
	./docx-comments $(ABS_TEST_IN) -o $(ABS_TEST_MD)
	./docx-comments $(ABS_TEST_MD) -o $(ABS_TEST_DOCX)
	@echo "Wrote $(TEST_MD) and $(TEST_DOCX) for manual inspection."

clean-roundtrip-example:
	rm -f $(TEST_MD) $(TEST_DOCX)
