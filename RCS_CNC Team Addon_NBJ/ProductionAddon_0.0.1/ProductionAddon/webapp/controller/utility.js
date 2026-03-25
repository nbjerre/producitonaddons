define([], function () {
  function StringJoin(arr, sep) {
    return arr.join(sep);
  }
  class Assert {
    static equal(actual, expected, message) {
      console.assert(actual == expected, message);
      if (actual != expected) {
        const details = `actual is '${actual}', expected should be '${expected}'`;
        throw new Error(message + " \n" + details);
      }
    }
    static notEqual(actual, expected, message) {
      console.assert(actual != expected, message);
      if (actual == expected) {
        const details = `actual is '${actual}', expected should not be '${expected}'`;
        throw new Error(message + "\n" + details);
      }
    }
  }
  return {
    StringJoin: StringJoin,
    Assert: Assert
  }
})