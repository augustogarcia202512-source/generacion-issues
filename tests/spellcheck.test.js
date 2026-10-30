import assert from 'node:assert/strict';
import dictionary from 'dictionary-es';
import nspell from 'nspell';

const spell = nspell(dictionary);

assert.equal(spell.correct('descripción'), true);
assert.equal(spell.correct('descripcion'), false);
assert.ok(spell.suggest('descripcion').includes('descripción'));
assert.equal(spell.correct('comntario'), false);
assert.ok(spell.suggest('comntario').includes('comentario'));
assert.equal(spell.correct('comentario'), true);
assert.ok(spell.suggest('comntario').includes('comentario'));

console.log('Spanish spell-check dictionary tests passed.');
