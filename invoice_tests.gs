function testNormalizeName() {
  var testCases = [
    {
      description: "Input with leading/trailing spaces",
      input: "  Test Name  ",
      expected: "Test-Name",
      to_upper_case: false
    },
    {
      description: "Input with diacritics",
      input: "Élève Test",
      expected: "Eleve-Test",
      to_upper_case: false
    },
    {
      description: "Input with special characters",
      input: "Test/Name_One.Two",
      expected: "Test-Name-One-Two",
      to_upper_case: false
    },
    {
      description: "Input for uppercase conversion",
      input: "test name",
      expected: "TEST-NAME",
      to_upper_case: true
    },
    {
      description: "Input with mixed special characters and spaces",
      input: "  Test/Name_One.Two  ",
      expected: "Test-Name-One-Two",
      to_upper_case: false
    },
    {
      description: "Input with diacritics and uppercase conversion",
      input: "Élève Test",
      expected: "ELEVE-TEST",
      to_upper_case: true
    },
    {
      description: "Input with numbers (should be removed)",
      input: "Test123Name456",
      expected: "TestName",
      to_upper_case: false
    },
    {
      description: "Input with multiple hyphens (should be reduced to one)",
      input: "Test---Name",
      expected: "Test-Name",
      to_upper_case: false
    },
    {
      description: "Input with trailing hyphen (should be removed)",
      input: "Test-Name-",
      expected: "Test-Name",
      to_upper_case: false
    }
  ];

  var failures = 0;

  for (var i = 0; i < testCases.length; i++) {
    var tc = testCases[i];
    var actual = normalizeName(tc.input, tc.to_upper_case);
    if (actual !== tc.expected) {
      failures++;
      Logger.log("--------------------------------------------------");
      Logger.log("FAIL: testNormalizeName - " + tc.description);
      Logger.log("Input: '" + tc.input + "' (to_upper_case: " + tc.to_upper_case + ")");
      Logger.log("Expected: '" + tc.expected + "'");
      Logger.log("Got: '" + actual + "'");
    }
  }
  
  if (failures > 0) {
    Logger.log("--------------------------------------------------");
    Logger.log("testNormalizeName: " + failures + " test(s) failed.");
  }
  // No summary log if all passed, just the "Finished" log.
  Logger.log("Finished testNormalizeName().");
  return failures === 0;
}

function testPlural() {
  var testCases = [
    {
      description: "Singular input",
      number: 1,
      message: "item",
      expected: "item"
    },
    {
      description: "Plural input",
      number: 2,
      message: "item",
      expected: "items"
    },
    {
      description: "Zero input",
      number: 0,
      message: "item",
      expected: "item"
    },
    {
      description: "Multi-word input",
      number: 3,
      message: "apple tree",
      expected: "apples trees"
    },
    {
      description: "Input with leading/trailing spaces, singular",
      number: 1,
      message: " apple tree ",
      expected: " apple tree "
    },
    {
      description: "Input with leading/trailing spaces, plural",
      number: 2,
      message: " apple tree ",
      expected: " apples trees "
    },
    {
      description: "Input with only spaces, singular",
      number: 1,
      message: " ",
      expected: " "
    },
    {
      description: "Input with only spaces, plural",
      number: 2,
      message: " ",
      expected: " "
    }
  ];

  var failures = 0;

  for (var i = 0; i < testCases.length; i++) {
    var tc = testCases[i];
    var actual = Plural(tc.number, tc.message);
    if (actual !== tc.expected) {
      failures++;
      Logger.log("--------------------------------------------------");
      Logger.log("FAIL: testPlural - " + tc.description);
      Logger.log("Input number: " + tc.number + ", Input message: '" + tc.message + "'");
      Logger.log("Expected: '" + tc.expected + "'");
      Logger.log("Got: '" + actual + "'");
    }
  }

  if (failures > 0) {
    Logger.log("--------------------------------------------------");
    Logger.log("testPlural: " + failures + " test(s) failed.");
  }
  // No summary log if all passed, just the "Finished" log.
  Logger.log("Finished testPlural().");
  return failures === 0;
}

function testGetDoBYear() {
  var testCases = [
    {
      description: "Date string (MM/DD/YYYY)",
      input: "12/31/1995",
      input_type: "string",
      expected: 1995
    },
    {
      description: "JavaScript Date object (Month is 0-indexed)",
      input: new Date(2005, 0, 1), // Month 0 is January
      input_type: "object",
      expected: 2005
    },
    {
      description: "Date string (YYYY-MM-DD)",
      input: "1987-05-15",
      input_type: "string",
      expected: 1987
    },
    {
      description: "Date string (M/D/YYYY without leading zeros)",
      input: "1/5/2003",
      input_type: "string",
      expected: 2003
    },
    {
      description: "JavaScript Date object (leap year)",
      input: new Date(2000, 1, 29), // Feb 29, 2000
      input_type: "object",
      expected: 2000
    }
  ];

  var failures = 0;

  for (var i = 0; i < testCases.length; i++) {
    var tc = testCases[i];
    var actual = getDoBYear(tc.input);
    if (actual !== tc.expected) {
      failures++;
      Logger.log("--------------------------------------------------");
      Logger.log("FAIL: testGetDoBYear - " + tc.description);
      var inputString = "";
      if (tc.input_type === "string") {
        inputString = "\"" + tc.input + "\"";
      } else {
        // For Date objects, construct a string representation
        inputString = "new Date(" + tc.input.getFullYear() + ", " + tc.input.getMonth() + ", " + tc.input.getDate() + ")";
      }
      Logger.log("Input: " + inputString);
      Logger.log("Expected: " + tc.expected);
      Logger.log("Got: " + actual);
    }
  }

  if (failures > 0) {
    Logger.log("--------------------------------------------------");
    Logger.log("testGetDoBYear: " + failures + " test(s) failed.");
  }
  // No summary log if all passed, just the "Finished" log.
  Logger.log("Finished testGetDoBYear().");
  return failures === 0;
}

function runInvoiceTests() {
  Logger.log("Starting all invoice tests...");
  var failedSuites = [];

  if (!testNormalizeName()) {
    failedSuites.push("testNormalizeName");
  }
  if (!testPlural()) {
    failedSuites.push("testPlural");
  }
  if (!testGetDoBYear()) {
    failedSuites.push("testGetDoBYear");
  }
  if (failedSuites.length === 0) {
    Logger.log("Summary: All test suites passed successfully!");
  } else {
    Logger.log("Summary: Failures detected in the following test suite(s): " + failedSuites.join(", ") + ".");
  }
}
