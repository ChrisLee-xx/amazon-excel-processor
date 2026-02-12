"""Product Name 规范化模块测试"""

import pytest
from amazon_excel_processor.name_normalizer import (
    collapse_spaces,
    format_variant_row,
    remove_numeric_suffix,
    replace_underscores,
    deduplicate_words,
)


class TestCollapseSpaces:
    def test_multiple_spaces(self):
        assert collapse_spaces("Royal Pegasus  Under   Moon") == "Royal Pegasus Under Moon"

    def test_leading_trailing(self):
        assert collapse_spaces("  Hello World  ") == "Hello World"

    def test_single_spaces_unchanged(self):
        assert collapse_spaces("A B C") == "A B C"


class TestFormatVariantRow:
    def test_frame_variant(self):
        name = (
            "Royal Pegasus Under Moon Canvas Print, Mythical Winged Horse Art-1 "
            "Frame-royal Pegasus Under Moon Canvas Print, Mythi08x12inch(20x30cm)"
        )
        result = format_variant_row(name)
        assert result == (
            "Royal Pegasus Under Moon Canvas Print, Mythical Winged Horse Art-1 "
            "Frame-style 08x12inch(20x30cm)"
        )

    def test_unframe_variant(self):
        name = (
            "Royal Pegasus Under Moon Canvas Print, Mythical Winged Horse Art-1 "
            "Unframe-royal Pegasus Under Moon Canvas Print, Mythi12x18inch(30x45cm)"
        )
        result = format_variant_row(name)
        assert result == (
            "Royal Pegasus Under Moon Canvas Print, Mythical Winged Horse Art-1 "
            "Unframe-style 12x18inch(30x45cm)"
        )

    def test_parent_row_unchanged(self):
        name = "Royal Pegasus Under Moon Canvas Print, Mythical Winged Horse Art-1"
        assert format_variant_row(name) == name

    def test_no_size_pattern_unchanged(self):
        name = "Some Title Frame-garbage without size"
        assert format_variant_row(name) == name


class TestRemoveNumericSuffix:
    def test_single_digit(self):
        assert remove_numeric_suffix("Mythical Winged Horse Art-1") == "Mythical Winged Horse Art"

    def test_multi_digit(self):
        assert remove_numeric_suffix("Canvas Print-12") == "Canvas Print"

    def test_no_suffix(self):
        assert remove_numeric_suffix("Canvas Print Art") == "Canvas Print Art"

    def test_preserve_frame_style(self):
        assert remove_numeric_suffix("Frame-style 08x12inch(20x30cm)") == "Frame-style 08x12inch(20x30cm)"

    def test_suffix_mid_string(self):
        assert remove_numeric_suffix("Art-1 Frame-style 08x12inch(20x30cm)") == "Art Frame-style 08x12inch(20x30cm)"


class TestReplaceUnderscores:
    def test_underscores(self):
        assert replace_underscores("Canvas_Print_Art") == "Canvas Print Art"

    def test_no_underscores(self):
        assert replace_underscores("Canvas Print") == "Canvas Print"


class TestDeduplicateWords:
    def test_word_three_times(self):
        assert deduplicate_words("Art Canvas Art Print Art Poster") == "Art Canvas Art Print Poster"

    def test_word_two_times_kept(self):
        assert deduplicate_words("Art Canvas Art Print") == "Art Canvas Art Print"

    def test_no_duplicates(self):
        assert deduplicate_words("Canvas Print Poster") == "Canvas Print Poster"

    def test_case_insensitive(self):
        assert deduplicate_words("art Canvas Art print art") == "art Canvas Art print"
