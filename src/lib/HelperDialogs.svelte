<script lang="ts">
    import { fade } from "svelte/transition";
    import type { HelperType } from "../Scripts/HelperType";
    import { cubicInOut } from "svelte/easing";
    import CellSpacing from "../assets/CellSpacing.jpg"
    import IndentNegative from "../assets/IndentNegative.jpg"
    import IntentPositive from "../assets/IndentPositive.jpg"
    import ListNumberFormat from "../assets/ListNumberFormat.jpg"
    import ListSecondLinePosition from "../assets/ListSecondLineTextPosition.jpg"
    import ListSecondLinePosition2 from "../assets/ListSecondLineTextPosition2.jpg"
    import ParagraphSameLine from "../assets/ParagraphSameLine.gif"
    import ParagraphSamePage from "../assets/ParagraphSamePage.gif"
    import Shading from "../assets/Shading.jpg"
    import ListTabPosition from "../assets/ListTabPosition.jpg"
    import ChartAxis2D from "../assets/ChartAxis2D.jpg"
    import ChartAxis3D from "../assets/ChartAxis3D.jpg"
    import ChartAddLabelEveryX from "../assets/ChartAddLabelEveryX.jpg"
    import ChartRotateLabelAxis from "../assets/ChartRotateLabelAxis.jpg"
    import ChartLegendShadow from "../assets/ChartLegendShadow.jpg"
    import ChartDataLabel from "../assets/ChartDataLabel.jpg"
    import ChartTitleShadow from "../assets/ChartTitleShadow.jpg"
    import ChartDistanceLength1 from "../assets/ChartDistanceLength1.jpg"
    import ChartDistanceLength2 from "../assets/ChartDistanceLength2.jpg"
    import ChartAngle0 from "../assets/ChartAngle0.jpg"
    import ChartAngle45 from "../assets/ChartAngle45.jpg"
    import ChartExplosion0 from "../assets/ChartExplosion0.jpg"
    import ChartExplosion15 from "../assets/ChartExplosion15.jpg"
    import { lang } from "../Scripts/Language";
    const {helperType, callback}: {helperType: HelperType, callback: () => void} = $props();
</script>

<div
    class="dialog"
    in:fade={{ duration: 400, easing: cubicInOut }}
    out:fade={{ duration: 400, easing: cubicInOut }}
>
    <div>
        {#if helperType === "CellSpacing"}
            <h2>{lang("Table cell spacing")}:</h2>
            <p>{lang("You can choose how much empty space there should be around each cell.")}<br><b>{lang("Value unit")}:</b> {lang("points")}</p><br>
            <img alt={lang("Table with some empty space around each cell")} src={CellSpacing}>
        {:else if helperType === "Indent"}
            <h2>{lang("Paragraph indent")}:</h2>
            <p>{lang("By changing the paragraph indent value, you can put some empty space before the first line. You can both put a positive value (image 1) or a negative value (image 2)")}.</p><br><b>{lang("Value unit")}:</b> {lang("points")}<br>
            <img alt={lang("Paragraph with indent positive applied: the text of the first line starts after if compared to the text of the other lines")} src={IntentPositive}><br>
            <img alt={lang("Paragraph with indent negative applied: the text of the first line starts before the text of the other lines")} src={IndentNegative}>
        {:else if helperType === "ListNumberFormat"}
            <h2>{lang("List number format")}:</h2>
            <p>{lang("You can change how the numbers or bullet points should be displayed, by adding some text before or after them. In the image example, the")} <code> - Example)</code> {lang("text has been added after the number")}.<br><b>{lang("Value unit")}:</b> {lang("points")}</p>
            <img alt={lang("Custom number format")} src={ListNumberFormat}>
        {:else if helperType === "ListSecondLine"}
            <h2>{lang("Second line position")}:</h2>
            <p>{lang("With this property, you can change the space before the start of every line after the first one. For example, in the first image this property was set to a normal value, while in the second one was set to")} <code>40</code>.<br><b>{lang("Value unit")}:</b> {lang("points")}</p>
            <img alt={lang("Normal second line position property")} src={ListSecondLinePosition}>
            <img alt={lang("Second line position set to 40: the text of the second and third line is positioned more at the right than the first line")} src={ListSecondLinePosition2}>   
        {:else if helperType === "ParagraphSameLine"}
            <h2>{lang("Keep paragraph on the same line")}:</h2>
            <p>{lang("As the title says, if this option is enabled, the paragraph will be put in the same page as the paragraph which follows it. Note that, in the video, the last two lines of the Lorem ipsum have been made a new line")}.</p><br>
            <img alt={lang("GIF where the entire paragraph goes to a new page")} src={ParagraphSameLine}>
        {:else if helperType === "ParagraphSamePage"}
            <h2>{lang("Keep paragraph on the same page")}:</h2>
            <p>{lang("If enabled, all the lines of the paragraph will be kept in the same page. So, if at least one line should go in a new page, all the paragraph will be put in the new page")}.</p><br>
            <img alt={lang("GIF where the second last paragraph goes to the new page with the last paragraph")} src={ParagraphSamePage}>
        {:else if helperType === "Shading"}
            <h2>{lang("Shading options")}:</h2>
            <p>{lang("With shading options, you can customize the background color of the paragraph. Here you can see an example, done by mixing red and yellow and by setting the shading type to")} <code>40%</code>.</p><br>
            <img alt={lang("Custom shading")} src={Shading}>
        {:else if helperType === "LevelOutline"}
            <h2>{lang("Level outline")}:</h2>
            <p>{lang("The level the paragraph should be displayed if a summary is generated, from 1 to 9. If it shouldn't be displayed,")} <code>BodyText</code>{lang(" should be selected")}.</p>
        {:else if helperType === "ListPosition"}
            <h2>{lang("Tab position")}:</h2>
            <p>{lang("Choose how much tabulation space should be added between the number/bullet point and the first character of the list. If you haven't set it, you might see a really high value. This is normal, but know that usually setting it of a few tens of points is more than enough (since setting it to a high value, like thousands of points, will cause Word to reject the style change")}.<br><b>{lang("Value unit")}:</b> {lang("points")}</p><br>
            <img alt={lang("First item on a list, with a lot of spcae between the number and the first character")} src={ListTabPosition}>
        {:else if helperType === "ExcelAxisTypes"}
            <h2>{lang("Axis types")}:</h2>
            <p>{lang("On Excel, there are two main chart axis types: the \"Value\" one, that usually is the one on the y axis, and the \"Category\" one, that usually is the one on the x axis. On 3D graphs, there might be also another axis (the \"Series\" axis), but in most cases you don't need to touch it.")}</p>
            <img src={ChartAxis2D} alt="2D Chart"><br>
            <img src={ChartAxis3D} alt="3D Chart">
        {:else if helperType === "ExcelAxisLabelEveryX"}
        <h2>{lang("Add label every")} <i>x</i> {lang("entries")}</h2>
        <p>{lang("You can choose to show the label of that axis every x entries. For example, in the image below this property was set to 2 in the \"Category\" axis")}.</p>
        <img src={ChartAddLabelEveryX} alt={lang("Chart where the bottom axis label is shown every two items")}>
        {:else if helperType === "ExcelAxisRotation"}
        <h2>{lang("Axis label rotation")}:</h2>
        <p>{lang("You can change the rotation of the axis labels. In this example, it was set to \"25\". Also negative values are accepted.")}</p>
        <img src={ChartRotateLabelAxis} alt={lang("Chart where the bottom axis label is slightly rotated")}>
        {:else if helperType === "ExcelAxisLegendShadow"}
        <h2>{lang("Legend shadow")}:</h2>
        <p>{lang("A shadow effect is added to the legend rectangle.")}</p>
        <img src={ChartLegendShadow} alt={lang("Legend with a shadow effect")}>
        {:else if helperType === "ExcelDataLabel"}
        <h2>{lang("Data labels")}:</h2>
        <p>{lang("Data labels are the labels before each point, that show its value.")}</p>
        <img src={ChartDataLabel} alt={lang("Data labels")}>
        {:else if helperType === "ExcelTitleShadow"}
        <h2>{lang("Title shadow")}:</h2>
        <p>{lang("Add a shadow to the title border.")}</p>
        <img src={ChartTitleShadow} alt={lang("Title shadow")}>
        {:else if helperType === "ExcelDistanceLength"}
        <h2>{lang("Gap width")}:</h2>
        <p>{lang("The space that is left between each series. If the number is high, more space will be left. In the first image, this value is set to \"200\", while in the second it's set to \"100\".")}</p>
        <img src={ChartDistanceLength1} alt={lang("More space between the columns (high number)")}><br>
        <img src={ChartDistanceLength2} alt={lang("Less space between the columns (low number)")}>
        {:else if helperType === "ExcelChartAngle"}
        <h2>{lang("Chart angle")}:</h2>
        <p>{lang("Change the angle where the chart starts. In the first image, the pie starts normally, while in the second it starts at 45°")}.</p>
        <img src={ChartAngle0} alt={`${lang("Chart angle")}: 0`}><br>
        <img src={ChartAngle45} alt={`${lang("Chart angle")}: 45`}>
        {:else if helperType === "ExcelChartExplosion"}
        <h2>{lang("Chart explosion")}:</h2>
        <p>{lang("Add more space in the center of the chart. In the first image, the \"Explosion\" property is set to 0, while in the second one to 15")}.</p>
        <img src={ChartExplosion0} alt={`${lang("Chart explosion")}: 0`}><br>
        <img src={ChartExplosion15} alt={`${lang("Chart explosion")}: 15`}>
        {/if}<br>
        <button onclick={() => callback()}>{lang("Close")}</button>
</div>
</div>

<style>
    img, video {
        width: 100%;
        height: auto;
        border: 1px solid var(--text);
        display: block;
        border-radius: 8px;
    }
</style>