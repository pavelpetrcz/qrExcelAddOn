// Jest setup: polyfills and Office/Excel minimal mocks
import 'whatwg-fetch';

// Provide a minimal Office mock so module initialization runs
global.Office = {
  onReady: (cb) => {
    try {
      cb && cb({ host: 'Excel' });
    } catch (e) {
      // ignore
    }
  },
  HostType: { Excel: 'Excel' }
};

// Minimal Excel.run mock to exercise code paths in tests
global.Excel = {
  run: jest.fn(async (cb) => {
    // Create a minimal context object similar to Excel.run
    const context = {
      workbook: {
        getSelectedRange: () => ({ load: () => {}, format: { fill: { color: null } }, address: 'A1' }),
        worksheets: {
          getActiveWorksheet: () => ({
            getRange: (addr) => ({
              load: () => {},
              left: 1,
              top: 2,
              width: 100,
              height: 100
            }),
            shapes: {
              addImage: (b64) => ({ left: 0, top: 0, width: 100, height: 100 })
            }
          })
        }
      },
      sync: async () => {}
    };
    return cb(context);
  })
};
