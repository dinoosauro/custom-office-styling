interface Props {
    font?: Excel.ChartFont,
    border?: Excel.ChartBorder,
    fill?: Excel.ChartFill
}

export default async function UpdateFormatProperties(source: Props, dest: Props, ctx: Excel.RequestContext) {
    for (const mainProp of ["font", "border"]) {
        if (typeof source[mainProp as "font"] !== "undefined" && typeof dest[mainProp as "font"] !== "undefined") {
            for (const item in source[mainProp as "font"]) {
                if (item.startsWith("_")) continue;
                if (typeof (source[mainProp as "font"] as Excel.ChartFont)[item as "size"] !== undefined && (source[mainProp as "font"] as Excel.ChartFont)[item as "size"] !== (dest[mainProp as "font"] as Excel.ChartFont)[item as "size"]) {
                    try {
                        (dest[mainProp as "font"] as Excel.ChartFont)[item as "size"] = (source[mainProp as "font"] as Excel.ChartFont)[item as "size"];
                        await ctx.sync();
                    } catch(ex) {
                        console.warn(mainProp, item, ex);
                    }
                }
            }
        }
    }
}