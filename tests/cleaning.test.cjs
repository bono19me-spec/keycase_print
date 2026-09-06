const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '../app.js'), 'utf8');
const names = ['resetDateGroups', 'getDateGroupKey', 'getDateGroups', 'getActiveDateGroup', 'getGroupDateText', 'setGroupDates', 'getRcInfo', 'getDatePickerConfigs', 'addCleaningDays', 'cleaningDateOffset',
  'formatCleaningDates', 'getPickerDateText', 'getCleaningInfo', 'parseMonthDayDate',
  'parseDateInfoDates', 'toggleSavedDate', 'sortMonthDayDates'];
function setup() {
  const context = vm.createContext({ Date });
  vm.runInContext(`let settings = { printStayInfo: true, printCleaningInfo: true,
    printRcInfo: true }; let records = []; let dateGroupOverrides = {}; let activeDateGroups = {};`, context);
  for (const name of names) {
    const start = source.indexOf(`function ${name}(`);
    const next = source.indexOf('\nfunction ', start + 1);
    vm.runInContext(source.slice(start, next < 0 ? undefined : next), context);
  }
  return (code) => vm.runInContext(code, context);
}
const room = (date, nights, custom = '') => `({ arrivalDate: new Date('${date}T00:00:00'), stayInfo: '${nights}泊', customCleaning: '${custom}' })`;
test('automatic dates are capped at four and exclude checkout', () => {
  const run = setup();
  for (const [nights, expected] of [[1, ''], [3, ''], [4, '清掃日：7/4'],
    [7, '清掃日：7/4、7/7'], [12, '清掃日：7/4、7/7、7/10'],
    [13, '清掃日：7/4、7/7、7/10、7/13'], [30, '清掃日：7/4、7/7、7/10、7/13']]) {
    assert.equal(run(`getCleaningInfo(${room('2026-07-01', nights)})`), expected);
  }
});
test('manual cleaning dates apply only to the selected arrival group', () => {
  const run = setup();
  run(`records = [${room('2026-10-01', 12)}, ${room('2026-10-05', 7)}];
    const config = getDatePickerConfigs()[0];
    activeDateGroups[config.textKey] = '2026-10-01';
    setGroupDates(config, ['10/6', '10/9']);`);
  assert.equal(run('getCleaningInfo(records[0])'), '清掃日：10/6、10/9');
  assert.equal(run('getCleaningInfo(records[1])'), '清掃日：10/8、10/11');
  run(`activeDateGroups[config.textKey] = '2026-10-05'; setGroupDates(config, ['10/4', '10/6', '10/12']);`);
  assert.equal(run('getCleaningInfo(records[1])'), '清掃日：10/6');
  assert.equal(run('getCleaningInfo(records[0])'), '清掃日：10/6、10/9');
  run(`setGroupDates(config, []);`);
  assert.equal(run('getCleaningInfo(records[1])'), '');
  run(`delete dateGroupOverrides[config.textKey]['2026-10-05'];`);
  assert.equal(run('getCleaningInfo(records[1])'), '清掃日：10/8、10/11');
});
test('room changes are independent per group and individual overrides take priority', () => {
  const run = setup();
  run(`records = [${room('2026-10-01', 30)}, ${room('2026-10-05', 30)}];
    const config = getDatePickerConfigs()[1];
    activeDateGroups[config.textKey] = '2026-10-01'; toggleSavedDate(config, '10/16');
    activeDateGroups[config.textKey] = '2026-10-05'; toggleSavedDate(config, '10/20');`);
  assert.equal(run('getRcInfo(records[0])'), '部屋変更：10/16');
  assert.equal(run('getRcInfo(records[1])'), '部屋変更：10/20');
  run(`toggleSavedDate(config, '10/21');`);
  assert.equal(run('getRcInfo(records[1])'), '部屋変更：10/21');
  run(`records[1].customRc = '部屋変更：10/22';`);
  assert.equal(run('getRcInfo(records[1])'), '部屋変更：10/22');
  run('resetDateGroups();');
  assert.equal(run('getRcInfo(records[0])'), '');
});
test('manual room dates are filtered by each stay', () => {
  const run = setup();
  assert.equal(run(`getCleaningInfo(${room('2026-07-01', 7, '清掃日：7/1、7/4、7/7、7/8、7/10')})`), '清掃日：7/4、7/7');
});
test('year rollover preserves chronological dates and editable offsets', () => {
  const run = setup();
  run(`records = [${room('2026-12-28', 13)}];`);
  assert.equal(run(`getPickerDateText(getDatePickerConfigs()[0])`), '清掃日：12/31、1/3、1/6、1/9');
  assert.equal(run(`getCleaningInfo(records[0])`), '清掃日：12/31、1/3、1/6、1/9');
  run(`toggleSavedDate({label: '清掃', textKey: 'cleaningInfoCustomText', maxDates: 6}, '1/4');`);
  assert.equal(run(`getCleaningInfo(${room('2026-12-28', 7)})`), '清掃日：12/31、1/3');
});
test('leap year and invalid calendar dates', () => {
  const run = setup();
  assert.equal(run(`getCleaningInfo(${room('2028-02-26', 7)})`), '清掃日：2/29、3/3');
  assert.equal(run(`getCleaningInfo(${room('2026-02-26', 7, '清掃日：2/30、3/1')})`), '清掃日：3/1');
});
test('empty schedule, disabled printing, missing dates and free text', () => {
  const run = setup();
  assert.equal(run(`getCleaningInfo({stayInfo: '7泊'})`), '');
  assert.equal(run(`getCleaningInfo(${room('2026-07-01', 7, '清掃不要')})`), '清掃不要');
  run(`records = [${room('2026-07-01', 7)}]; setGroupDates(getDatePickerConfigs()[0], []);`);
  assert.equal(run(`getCleaningInfo(${room('2026-07-01', 7)})`), '');
  run('settings.printCleaningInfo = false');
  assert.equal(run(`getCleaningInfo(${room('2026-07-01', 7, '清掃日：7/4')})`), '');
});

test('groups combine equal check-in dates and distinguish years and missing dates', () => {
  const run = setup();
  run(`records = [${room('2026-10-01', 12)}, ${room('2026-10-01', 7)}, ${room('2026-10-05', 7)}, ${room('2027-10-01', 7)}, {stayInfo: ''}];`);
  assert.equal(run('getDateGroups().length'), 4);
  assert.equal(run('getDateGroups()[0].count'), 2);
  assert.equal(run('getDateGroups()[3].key'), 'unknown');
});
