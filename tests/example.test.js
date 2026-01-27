// Example test file
const assert = require('assert');

describe('Example Tests', () => {
  it('should pass basic test', () => {
    assert.equal(1 + 1, 2);
  });

  it('should handle async operations', async () => {
    const result = await Promise.resolve(true);
    assert.equal(result, true);
  });
});
