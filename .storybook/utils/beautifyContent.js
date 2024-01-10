import * as beautify from 'js-beautify';

export const beautifyContent = (type, content) => {
  const options = {
    indent_size: '2',
    indent_char: ' ',
    max_preserve_newlines: '0',
    preserve_newlines: true,
    keep_array_indentation: true,
    break_chained_methods: false,
    indent_scripts: 'separate',
    brace_style: 'collapse,preserve-inline',
    space_before_conditional: false,
    unescape_strings: false,
    jslint_happy: false,
    end_with_newline: true,
    wrap_line_length: '120',
    indent_inner_html: true,
    comma_first: false,
    e4x: true,
    indent_empty_lines: false
  };

  let beautifiedContent = content;
  switch (type) {
    case 'html':
      beautifiedContent = beautify.default.html(content, options);
      break;
    case 'js':
    case 'json':
      beautifiedContent = beautify.default.js(content, options);
      break;
    case 'css':
      beautifiedContent = beautify.default.css(content, options);
      break;
    default:
      break;
  }
  return beautifiedContent;
};
