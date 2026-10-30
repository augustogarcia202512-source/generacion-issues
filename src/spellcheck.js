import aff from '../node_modules/dictionary-es/index.aff';
import dic from '../node_modules/dictionary-es/index.dic';
import nspell from 'nspell';

const spell = nspell({aff, dic});

window.issueSpellcheck = {
  check(word) {
    return spell.correct(word);
  },
  suggest(word) {
    return spell.suggest(word).slice(0, 3);
  },
  add(word) {
    spell.add(word);
  }
};
