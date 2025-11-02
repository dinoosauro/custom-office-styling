<script lang="ts">
    import Card from "../../lib/Card.svelte";
    import { lang } from "../../Scripts/Language";
    import FontChange from "./FontChange.svelte";

    const {contentControl}: {contentControl: Word.ContentControl} = $props();
    /**
   * A Spinner element at the center of the page
   */
  const spinner = document.createElement("div");
  spinner.classList.add("spinner");

  // Variables used while adding a new dropdown item
  let displayedText = "";
  let value = "";
  let addIndex = -1;

  // Variables used while changing a dropdown item
  let selectedIndex = $state(1);
  let deleteSelectedItem = false;
  // @ts-ignore
  if (Array.isArray(contentControl.comboBoxContentControl?.listItems)) contentControl.comboBoxContentControl.listItems = {items: contentControl.comboBoxContentControl.listItems};
  // @ts-ignore
  if (Array.isArray(contentControl.dropDownListContentControl?.listItems)) contentControl.dropDownListContentControl.listItems = {items: contentControl.dropDownListContentControl.listItems};

</script>

<h2>{lang("Customize content control")}:</h2>
<Card secondCard={true}>
<h3>{lang("Appearance")}:</h3>
<label class="flex hcenter gap">
    {lang("Appearance")}: <div class="selectContainer">
        <select bind:value={contentControl.appearance}>
        {#each ["BoundingBox", "Tags", "Hidden"] as option}
        <option value={option}>{option}</option>
        {/each}
        </select>
    </div>
</label><br>
<label class="flex hcenter gap">
    {lang("Color")}: <input type="color" bind:value={contentControl.color}>
</label><br>
<Card>
    <h4>{lang("Font")}:</h4>
    <FontChange sourceFont={contentControl.font}></FontChange>
</Card>
</Card><br>
<Card secondCard={true}>
    <h3>{lang("Permissions")}:</h3>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={contentControl.cannotDelete}>
        {lang("Cannot delete")}
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={contentControl.cannotDelete}>
        {lang("Cannot edit")}
    </label><br>
</Card><br>
{#if contentControl.type === "ComboBox" || contentControl.type === "DropDownList"}
<Card secondCard={true}>
    <h3>{lang("Item list")}:</h3>
    <Card>
        <h4>{lang("Add new item")}:</h4>
        <label class="flex hcenter gap">
            {lang("Displayed text")}: <input type="text" bind:value={displayedText}>
        </label><br>
        <label class="flex hcenter gap">
            {lang("Value")}: <input type="text" bind:value={value}>
        </label><br>
        <label class="flex hcenter gap">
            {lang("Index (negative value to put it in last position)")}: <input type="number" bind:value={addIndex}>
        </label>
        <button onclick={() => {
            document.body.append(spinner);
            setTimeout(() => {
                Word.run(async (ctx) => {
                    const controls = ctx.document.getSelection().getContentControls().load({$all: true, font: {$all: true}});
                    await ctx.sync();
                    for (const control of controls.items) {
                        if (control.type === "ComboBox" || control.type == "DropDownList") {
                            control.dropDownListContentControl
                            control[control.type === "ComboBox" ? "comboBoxContentControl" : "dropDownListContentControl"].addListItem(displayedText, value, addIndex < 0 ? undefined : addIndex);
                            await ctx.sync();
                        }
                    }
                    spinner.remove();
                })
            }, 1)
        }}>{lang("Add item to list")}</button>
    </Card>
    {#if (contentControl[contentControl.type === "ComboBox" ? "comboBoxContentControl" : "dropDownListContentControl"]?.listItems?.items?.length ?? 0) !== 0}
        <br>
        <Card>
            <h4>{lang("Edit single item")}:</h4>
            <label class="flex hcenter gap">
                {lang("Edit item in position")} <input min="1" type="number" bind:value={selectedIndex}>
            </label><br>
            <label class="flex hcenter gap">
                {lang("Title")}: <input type="text" bind:value={contentControl[contentControl.type === "ComboBox" ? "comboBoxContentControl" : "dropDownListContentControl"].listItems.items[selectedIndex].displayText}>
            </label><br>
            <label class="flex hcenter gap">
                {lang("Value")}: <input type="text" bind:value={contentControl[contentControl.type === "ComboBox" ? "comboBoxContentControl" : "dropDownListContentControl"].listItems.items[selectedIndex].value}>
            </label><br>
            <label class="flex hcenter gap">
                {lang("Index")}: <input type="number" bind:value={contentControl[contentControl.type === "ComboBox" ? "comboBoxContentControl" : "dropDownListContentControl"].listItems.items[selectedIndex].index}>
            </label><br>
            <label class="flex hcenter gap">
                <input type="checkbox" bind:checked={deleteSelectedItem}>{lang("Delete entry")}
            </label>
            <button onclick={() => {
                document.body.append(spinner);
                setTimeout(() => {
                    Word.run(async (ctx) => {
                        const controls = ctx.document.getSelection().getContentControls().load({$all: true, font: {$all: true}});
                        await ctx.sync();
                        for (const control of controls.items) {
                            if (control.type === "ComboBox" || control.type === "DropDownList") {
                                const items = control[control.type === "ComboBox" ? "comboBoxContentControl" : "dropDownListContentControl"].listItems.load();
                                await ctx.sync();
                                if (items.items.length >= selectedIndex) {
                                    const item = items.items[selectedIndex - 1];
                                    if (deleteSelectedItem) item.delete(); else {
                                        for (const prop of ["displayText", "value", "index"]) {
                                            if (contentControl[contentControl.type === "ComboBox" ? "comboBoxContentControl" : "dropDownListContentControl"].listItems.items[selectedIndex][prop as "index"] !== item[prop as "index"]) item[prop as "index"] = contentControl[contentControl.type === "ComboBox" ? "comboBoxContentControl" : "dropDownListContentControl"].listItems.items[selectedIndex].index;
                                        }
                                    }
                                    await ctx.sync();
                                }
                            }
                        }
                        spinner.remove();
                    })
                }, 1);
            }}>{lang("Change properties")}</button>
        </Card>
    {/if}
</Card><br>
{/if}
<Card secondCard={true}>
    <h3>{lang("Other settings")}:</h3>
    <label class="flex hcenter gap">
        {lang("Tag")}: <input type="text" bind:value={contentControl.tag}>
    </label><br>
    <label class="flex hcenter gap">
        {lang("Title")}: <input type="text" bind:value={contentControl.title}>
    </label>
</Card><br>
<button onclick={() => {
    document.body.append(spinner);
    setTimeout(() => {
        Word.run(async (ctx) => {
            const controls = ctx.document.getSelection().getContentControls().load({$all: true, font: {$all: true}});
            await ctx.sync();
            for (const control of controls.items) {
                await ctx.sync();
                for (const prop in contentControl) {
                    if (prop.startsWith("_") || prop === "id" || prop === "text" || typeof control[prop as "title"] === "object" || control[prop as "title"] === contentControl[prop as "title"]) continue;
                    try {
                        control[prop as "title"] = contentControl[prop as "title"];
                    } catch(ex) {
                        console.warn(prop, ex);
                    }
                }
                control.font.load({$all: true});
                await ctx.sync();
                for (const prop in control.font) {
                    try {
                        if (typeof control.font[prop as "size"] === "object" || control.font[prop as "size"] === contentControl.font[prop as "size"] || prop.startsWith("_")) continue;
                        control.font[prop as "size"] = contentControl.font[prop as "size"];
                    } catch(ex) {
                        console.warn(ex);
                    }
                }
                await ctx.sync();
            }
            spinner.remove();
        })
    }, 1)
}}>{lang("Apply")}</button>