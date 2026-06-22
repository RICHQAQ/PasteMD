from pastemd.utils.md_normalizer import fence_plain_text_code


def test_fence_plain_text_code_wraps_copied_code_snippet():
    text = "function greet(name) {\n  return `hello ${name}`;\n}"

    assert fence_plain_text_code(text) == "```\nfunction greet(name) {\n  return `hello ${name}`;\n}\n```"


def test_fence_plain_text_code_leaves_normal_markdown_alone():
    text = "# Title\n\nA normal paragraph\n\n- item"

    assert fence_plain_text_code(text) == text


def test_fence_plain_text_code_leaves_existing_fences_alone():
    text = "```js\nconsole.log('ok')\n```"

    assert fence_plain_text_code(text) == text
