let selectedId = "";
export async function GetChart(ctx: Excel.RequestContext) {
    let tempData = ctx.workbook.worksheets.getActiveWorksheet().charts.load();
    await ctx.sync();
    return tempData.items.find(i => i.id === selectedId) as Excel.Chart;
}
export async function UpdateSelectedChart(id: string) {
    selectedId = id;
}