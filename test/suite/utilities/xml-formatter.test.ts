import { expect } from 'chai';
import { XmlFormatter } from '../../../src/utilities/xml-formatter';

suite('OOXMLViewer Xml Formatter', function () {
  var isXmlTests = [
    {
      description: 'contents are xml',
      content: `<?xml ?><Types></Types>`,
      success: true,
    },
    {
      description: "string that doesn't start with <?xml",
      content: `notxml`,
      success: false,
    },
    {
      description: 'contents are binary',
      content: `\x00\x01\x02\x03\x04`,
      success: false,
    },
  ];

  isXmlTests.forEach(function (args) {
    test(`isXml where ${args.description} returns ${args.success}`, function () {
      const encoder = new TextEncoder();
      const data = encoder.encode(args.content);

      const equal = XmlFormatter.isXml(data);

      expect(equal).to.equal(args.success);
    });
  });

  var areEqualTests = [
    {
      description: 'contents are the same',
      content1: `<?xml ?>
<Types>
</Types>`,
      content2: `<?xml ?>
<Types>
</Types>`,
      success: true,
    },
    {
      description: 'contents are the same when minified',
      content1: `<?xml ?><Types></Types>`,
      content2: `<?xml ?>
<Types>
</Types>`,
      success: true,
    },
    {
      description: 'contents are not the same when minified',
      content1: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
      content2: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
</Types>`,
      success: false,
    },
  ];

  areEqualTests.forEach(function (args) {
    test(`areEqual where ${args.description} returns ${args.success}`, function () {
      const encoder = new TextEncoder();
      const data1 = encoder.encode(args.content1);
      const data2 = encoder.encode(args.content2);

      const equal = XmlFormatter.areEqual(data1, data2);

      expect(equal).to.equal(args.success);
    });
  });

  var minifyTests = [
    {
      preserveComments: true,
      raw: `<?xml ?><!-- comment -->
<Types>
</Types>`,
      minified: `<?xml ?><!-- comment --><Types></Types>`,
    },
    {
      preserveComments: false,
      raw: `<?xml ?><!-- comment -->
<Types>
</Types>`,
      minified: `<?xml ?><Types></Types>`,
    },
  ];

  minifyTests.forEach(function (args) {
    test(`minify when preserve comments is ${args.preserveComments} where ${args.raw} returns ${args.minified}`, function () {
      const encoder = new TextEncoder();
      const decoder = new TextDecoder();
      const rawData = encoder.encode(args.raw);

      const minified = XmlFormatter.minify(rawData, args.preserveComments);

      expect(decoder.decode(minified)).to.equal(args.minified);
    });
  });

  var formatTests = [
    {
      raw: `<?xml ?><!-- comment --><Types></Types>`,
      formatted: `<?xml?>\r\n<!-- comment -->\r\n<Types></Types>`,
    },
  ];

  formatTests.forEach(function (args) {
    test(`format where ${args.raw} returns ${args.formatted}`, function () {
      const encoder = new TextEncoder();
      const decoder = new TextDecoder();
      const rawData = encoder.encode(args.raw);

      const formatted = XmlFormatter.format(rawData);

      expect(decoder.decode(formatted)).to.equal(args.formatted);
    });
  });

  test(`format where not xml returns original content`, function () {
    const encoder = new TextEncoder();
    const decoder = new TextDecoder();
    const rawData = encoder.encode('not-xml<>\r\n<>');

    const formatted = XmlFormatter.format(rawData);

    expect(decoder.decode(formatted)).to.equal('not-xml<>\r\n<>');
  });

  test(`minify where not xml returns original content`, function () {
    const encoder = new TextEncoder();
    const decoder = new TextDecoder();
    const rawData = encoder.encode('not-xml<>\r\n<>');

    const formatted = XmlFormatter.minify(rawData, false);

    expect(decoder.decode(formatted)).to.equal('not-xml<>\r\n<>');
  });
});
