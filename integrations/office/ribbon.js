// Commands share the task pane runtime: no second client, polling loop, URL
// navigation or persisted command channel may replace a live editing session.
export function registerRibbon({ Office, ready, insert, edit, report }) {
  if (!Office.actions?.associate) return;
  const command = (run) => async (event) => {
    try {
      await ready;
      await Office.addin.showAsTaskpane();
      await run();
    } catch (error) {
      report(error.message);
    } finally {
      event.completed();
    }
  };
  Office.actions.associate("reshikiInsert", command(insert));
  Office.actions.associate("reshikiEdit", command(edit));
}
