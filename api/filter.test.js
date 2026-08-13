import test from 'node:test';
import assert from 'node:assert/strict';
import { buildGraphFilter } from './index.js';

test('buildGraphFilter supports positive and negative sender/subject filters', () => {
  const filter = buildGraphFilter({
    sender: 'alerts@',
    subject: 'invoice',
    senderExclude: 'promo@',
    subjectExclude: 'newsletter',
  });

  assert.match(filter, /contains\(from\/emailAddress\/address,'alerts@'/);
  assert.match(filter, /contains\(subject,'invoice'/);
  assert.match(filter, /not\(contains\(from\/emailAddress\/address,'promo@'\)\)/);
  assert.match(filter, /not\(contains\(subject,'newsletter'\)\)/);
});
