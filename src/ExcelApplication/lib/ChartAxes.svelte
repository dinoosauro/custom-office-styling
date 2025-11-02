<script lang="ts">
    import Card from "../../lib/Card.svelte";
    import HelperDialogs from "../../lib/HelperDialogs.svelte";
    import type { HelperType } from "../../Scripts/HelperType";
    import { lang } from "../../Scripts/Language";
    import {GetChart} from "../Scripts/GetChart";
    import GridlineStyling from "./GridlineStyling.svelte";

    let selectedAxis = $state("categoryAxis");
    let helper = $state<HelperType | undefined>(undefined);
    let {data, spinner}: {data: Excel.Chart, spinner: HTMLDivElement} = $props();
</script>

<h2>{lang("Axis")}:</h2>
        <label class="flex hcenter gap">
            <span class="help" onclick={() => (helper = "ExcelAxisTypes")}>{lang("Edit the following axis")}:</span> <div class="selectContainer"><select bind:value={selectedAxis}>
                <option value="valueAxis">{lang("Value")}</option>
                <option value="categoryAxis">{lang("Category")}</option>
                <option value="seriesAxis">{lang("Series")}</option>
            </select></div>
        </label><br />
        <Card secondCard={true}>
            <h3>{lang("Styling")}:</h3>
            <label class="flex hcenter gap">
                {lang("Alignment")}: <div class="selectContainer"><select>
                    {#each ["Left", "Center", "Right"] as option}
                        <option value={option}>{lang(option)}</option>
                    {/each}
                </select></div>
            </label><br />
            <label class="flex hcenter gap">
                {lang("Height (in points)")}: <input
                    type="number"
                    bind:value={
                        data.axes[selectedAxis as "categoryAxis"].height
                    }
                />
            </label><br />
            <label class="flex hcenter gap">
                {lang("Distance from left")}: <input
                    type="number"
                    bind:value={data.axes[selectedAxis as "categoryAxis"].left}
                />
            </label><br />
        </Card><br />
        <Card secondCard={true}>
            <h3>{lang("Chart range")}:</h3>
            <label class="flex hcenter gap">
                {lang("Minimum")}: <input
                    type="number"
                    bind:value={
                        data.axes[selectedAxis as "categoryAxis"].minimum
                    }
                />
            </label><br />
            <label class="flex hcenter gap">
                {lang("Maximum")}: <input
                    type="number"
                    bind:value={
                        data.axes[selectedAxis as "categoryAxis"].maximum
                    }
                />
            </label><br />
            <label class="flex hcenter gap">
                {lang("Number format")}: <input
                    bind:value={
                        data.axes[selectedAxis as "categoryAxis"].numberFormat
                    }
                />
            </label>
        </Card><br />
        <Card secondCard={true}>
            <h3>{lang("Major ticks")}:</h3>
                <label class="flex hcenter gap">
                    {lang("Tick mark")}: <div class="selectContainer"> <select
                        bind:value={
                            data.axes[selectedAxis as "categoryAxis"]
                                .majorTickMark
                        }
                    >
                        {#each ["Cross", "Inside", "Outside"] as option}
                            <option value={option}>{lang(option)}</option>
                        {/each}
                    </select></div>
                </label><br />
                <label class="flex hcenter gap">
                    {lang("Interval between tick marks")}: <input
                        type="number"
                        bind:value={
                            data.axes[selectedAxis as "categoryAxis"].majorUnit
                        }
                    />
                </label><br />
                <Card secondCard={false}>
                    <h4>{lang("Styling")}:</h4>
                    <GridlineStyling
                        gridlineFormat={data.axes[
                            selectedAxis as "categoryAxis"
                        ].majorGridlines}
                    ></GridlineStyling>
                </Card>
        </Card><br />
        <Card secondCard={true}>
            <h3>{lang("Minor ticks")}:</h3>
                <label class="flex hcenter gap">
                    {lang("Tick mark")}: <div class="selectContainer"> <select
                        bind:value={
                            data.axes[selectedAxis as "categoryAxis"]
                                .minorTickMark
                        }
                    >
                        {#each ["Cross", "Inside", "Outside"] as option}
                            <option value={option}>{lang(option)}</option>
                        {/each}
                    </select></div>
                </label><br />
                <label class="flex hcenter gap">
                    {lang("Interval between tick marks")}: <input
                        type="number"
                        bind:value={
                            data.axes[selectedAxis as "categoryAxis"].minorUnit
                        }
                    />
                </label><br />
                <Card secondCard={false}>
                    <h4>{lang("Styling")}:</h4>
                    <GridlineStyling
                        gridlineFormat={data.axes[
                            selectedAxis as "categoryAxis"
                        ].minorGridlines}
                    ></GridlineStyling>
                </Card>
        </Card><br />
        <Card secondCard={true}>
            <h3>{lang("Ticks")}:</h3>
            <label class="flex hcenter gap">
                <input type="checkbox" bind:checked={data.axes[selectedAxis as "categoryAxis"].showDisplayUnitLabel}>
                {lang("Show unit label")} 
            </label><br>
            <label class="flex hcenter gap">
                <span class="help" onclick={() => (helper = "ExcelAxisRotation")}>{lang("Tick mark label rotation")}:</span> <input type="number" bind:value={data.axes[selectedAxis as "categoryAxis"].textOrientation}>
            </label><br>
            <label class="flex hcenter gap">
                {lang("Position of tick label")} <div class="selectContainer"><select bind:value={data.axes[selectedAxis as "categoryAxis"].tickLabelPosition}>
                    <option value="NextToAxis">{lang("Next to the axis")}</option>
                    <option value="High">{lang("High")}</option>
                    <option value="Low">{lang("Low")}</option>
                    <option value="None">{lang("None")}</option>
                </select></div>
            </label><br>
            <label class="flex hcenter gap">
                {lang("Spacing between tick marks")}: <input type="number" bind:value={data.axes[selectedAxis as "categoryAxis"].tickMarkSpacing}>
            </label><br>
            <label class="flex hcenter gap">
                <span class="help" onclick={() => (helper = "ExcelAxisLabelEveryX")}>{lang("Add a label every (this) categories/values")}:</span> <input type="number" bind:value={data.axes[selectedAxis as "categoryAxis"].tickLabelSpacing}>
            </label><br>
        </Card><br>
        <Card secondCard={true}>
            <h3>{lang("Advanced")}:</h3>
            <label class="flex hcenter gap">
                {lang("Axis category type")}: <div class="selectContainer"><select
                    bind:value={
                        data.axes[selectedAxis as "categoryAxis"].alignment
                    }
                >
                    <option value="Automatic">{lang("Automatic")}</option>
                    <option value="TextAxis">{lang("Text")}</option>
                    <option value="DateAxis">{lang("Date")}</option>
                </select></div>
            </label><br>
            <label class="flex hcenter gap">
                {lang("Base time unit")}: <div class="selectContainer"> <select
                    bind:value={
                        data.axes[selectedAxis as "categoryAxis"].baseTimeUnit
                    }
                >
                    {#each ["Days", "Months", "Years"] as option}
                        <option value={option}>{lang(option)}</option>
                    {/each}
                </select></div>
            </label><br />
            <label class="flex hcenter gap">
                {lang("Display unit")}: <input
                    type="number"
                    bind:value={
                        data.axes[selectedAxis as "categoryAxis"]
                            .customDisplayUnit
                    }
                />
            </label><br />
        </Card><br>
        <Card secondCard={true}>
            <h3>{lang("Scale type")}:</h3>
            <label class="flex hcenter gap">
                {lang("Scale type")}: <div class="selectContainer"><select bind:value={data.axes[selectedAxis as "categoryAxis"]}>
                    <option value="Linear">{lang("Linear")}</option>
                    <option value="Logarithmic">{lang("Logarithmic")}</option>
                </select></div>
            </label><br>
            <label class="flex hcenter gap">
                {lang("Logarithmic base")}: <input
                    type="number"
                    bind:value={data.axes[selectedAxis as "categoryAxis"].logBase}
                />
            </label>
        </Card><br>
        <button
            onclick={() => {
                document.body.append(spinner);
                setTimeout(() => {
                    Excel.run(async (ctx) => {
                        const chart = await GetChart(ctx);
                        const axes = chart.axes.load();
                        const currentAxis =
                            axes[selectedAxis as "categoryAxis"].load();
                            await ctx.sync();
                            // We'll separately try to get the major and the minor gridline to avoid errors if one of them is not available
                        try {
                            currentAxis.majorGridlines.load({$all: true, format: {$all: true, line: {$all: true}}});
                            await ctx.sync();
                        } catch (ex) {
                            console.warn(ex);
                        }
                        try {
                            currentAxis.minorGridlines.load({$all: true, format: {$all: true, line: {$all: true}}});
                            await ctx.sync();
                        } catch (ex) {
                            console.warn(ex);
                        }
                        // Let's start copying properties, only if they are different from the ones already applied to the chart.
                        for (const prop in data?.axes[
                            selectedAxis as "categoryAxis"
                        ]) {
                            try {
                                if (
                                    data?.axes[selectedAxis as "categoryAxis"][
                                        prop as "numberFormat"
                                    ] === currentAxis[prop as "numberFormat"]
                                )
                                    continue;
                                if (
                                    typeof currentAxis[prop as "numberFormat"] ===
                                    "object"
                                ) {
                                    for (const secondProp in data?.axes[
                                        selectedAxis as "categoryAxis"
                                    ][prop as "majorGridlines"]) {
                                        try {
                                            if (
                                                typeof data?.axes[
                                                    selectedAxis as "categoryAxis"
                                                ][prop as "majorGridlines"][
                                                    secondProp as "format"
                                                ] === "object"
                                            ) {
                                                // NOTE: Currently, this is called only for the gridline. Therefore, the "line" object is hardcoded.
                                                const dataToCompare = JSON.parse(JSON.stringify(currentAxis[
                                                                prop as "majorGridlines"
                                                            ][
                                                                secondProp as "format"
                                                            ].line));
                                                for (const thirdProp in data.axes[selectedAxis as "categoryAxis"][prop as "majorGridlines"][secondProp as "format"].line) {
                                                    try {
                                                        try {
                                                            if (
                                                                data.axes[selectedAxis as "categoryAxis"][prop as "majorGridlines"][secondProp as "format"].line[
                                                                    thirdProp as "color"
                                                                ] ===
                                                                dataToCompare[
                                                                    thirdProp as "color"
                                                                ]
                                                            ) continue;
                                                        } catch(ex) {
                                                            
                                                        }
                                                            currentAxis[
                                                                prop as "majorGridlines"
                                                            ][
                                                                secondProp as "format"
                                                            ].line[
                                                                thirdProp as "color"
                                                            ] = data.axes[
                                                                selectedAxis as "categoryAxis"
                                                            ][
                                                                prop as "majorGridlines"
                                                            ][
                                                                secondProp as "format"
                                                            ].line[
                                                                thirdProp as "color"
                                                            ];
                                                        await ctx.sync();
                                                    } catch (ex) {
                                                        console.warn(
                                                            prop,
                                                            secondProp,
                                                            thirdProp,
                                                            ex,
                                                        );
                                                    }
                                                }
                                            if (
                                                data?.axes[
                                                    selectedAxis as "categoryAxis"
                                                ][prop as "majorGridlines"][
                                                    secondProp as "visible"
                                                ] ===
                                                currentAxis[
                                                    prop as "majorGridlines"
                                                ][secondProp as "visible"]
                                            )
                                                continue;
                                                
                                                currentAxis[
                                                    prop as "majorGridlines"
                                                ][secondProp as "visible"] = data.axes[
                                                    selectedAxis as "categoryAxis"
                                                ][prop as "majorGridlines"][
                                                    secondProp as "visible"
                                                ]
                                                await ctx.sync();
                                            }
                                        } catch (ex) {
                                            console.warn(prop, secondProp, ex);
                                        }
                                    }
                                }
                                if (prop === "customDisplayUnit")
                                    currentAxis.setCustomDisplayUnit(
                                        data.axes[selectedAxis as "categoryAxis"]
                                            .customDisplayUnit,
                                    );
                                else
                                    currentAxis[prop as "numberFormat"] =
                                        data.axes[selectedAxis as "categoryAxis"][
                                            prop as "numberFormat"
                                        ];
                                        await ctx.sync();
                            } catch (ex) {
                                console.warn(prop, ex);
                            }
                        }
                        await ctx.sync();
                        spinner.remove();
                    });
                }, 1)
            }}>{lang("Apply")}</button
        >
{#if helper}
  <HelperDialogs helperType={helper} callback={() => (helper = undefined)}></HelperDialogs>
{/if}
