<script lang="ts">
    import { onMount } from "svelte";
    import Excel from "./ExcelApplication/Excel.svelte";
    import Word from "./WordApplication/Word.svelte";
    import { lang, updateOfficeReady } from "./Scripts/Language";
    import { fade } from "svelte/transition";
    import { cubicInOut } from "svelte/easing";
    import OpenSourceDialog from "./lib/OpenSourceDialog.svelte";

    /**
     * If the dialog to go back to the home should be triggered.
     */
    let showHomeDialog = $state(false);
    /**
     * The link that needs to be displayed in the "Download file" dialog.
     * It's false if the dialog doesn't need to be displayed, either because the user hasn't asked to download a file or because the OpenBrowserWindowApi is supported.
     */
    let downloadLink: false | string = $state(false);

    /**
     * The icon of the website, dynamically updated when the user switches from light to dark mode and viceversa
     */
    let titleImage: HTMLImageElement;
    /**
     * If true, the dialog with the licenses will be shown
     */
    let showLicenseDialog = $state(false);
    let errorDialog: false | string = $state(false);
    let isFromExcel = $state(false);

    // Dark/light mode change part

    /**
     * CSS properties that should be changed from light to dark mode
     */
    const itemsToChange = [
        "text",
        "background",
        "card",
        "input",
        "accent",
        "hover-filter",
        "active-filter",
    ];
    const darkThemeVariant = itemsToChange.map((i) =>
        getComputedStyle(document.body).getPropertyValue(`--${i}`),
    );
    const lightThemeVariant = [
        "#151515",
        "#fafafa",
        "#d2d2d2",
        "#b4b4b4",
        "#bbc698",
        "85%",
        "75%",
    ];
    let forceReRender = $state(`Entire-${Date.now()}`);
    let isLightTheme = false;
    /**
     * Switch between the light and the dark theme
     * @param skipLocalStorageSaving if true, the preference won't be saved in the LocalStorage. This should be done in the case the preference is fetched from the Word UI, and not from user input.
     */
    function changeTheme(skipLocalStorageSaving?: boolean) {
        isLightTheme = !isLightTheme;
        titleImage.src = isLightTheme ? "./logo_dark.svg" : "./logo_light.svg";
        for (let i = 0; i < itemsToChange.length; i++) {
            document.body.style.setProperty(
                `--${itemsToChange[i]}`,
                isLightTheme ? lightThemeVariant[i] : darkThemeVariant[i],
            );
        }
        !skipLocalStorageSaving &&
            localStorage.setItem(
                "CustomWordStyle-Theme",
                isLightTheme ? "light" : "dark",
            );
    }
    onMount(() => {
        console.error = (err) => (errorDialog = err.toString());
        window.addEventListener("error", (e) => {
            errorDialog = `${e.error.toString()}\n\nFrom line ${e.lineno}, column ${e.colno} of ${e.filename}`;
            document.querySelector(".spinner")?.remove();
        });
        window.addEventListener("unhandledrejection", (e) => {
            errorDialog = e.reason;
            document.querySelector(".spinner")?.remove();
        });
        Office.onReady().then(() => {
            updateOfficeReady();
            if (
                (!Office.context.officeTheme.isDarkTheme &&
                    localStorage.getItem("CustomWordStyle-Theme") !== "dark") ||
                localStorage.getItem("CustomWordStyle-Theme") === "light"
            ) {
                changeTheme(true);
            }
            isFromExcel = Office.context.host === Office.HostType.Excel;
        });
        setTimeout(() => (forceReRender = `Entire-${Date.now()}`), 150);
    });
</script>

{#key forceReRender}
    <header>
        <div class="flex hcenter gap">
            <img
                bind:this={titleImage}
                onclick={() => {
                    showHomeDialog = true;
                }}
                class="hover"
                style="width: 48px; height: 48px"
                src="./logo_light.svg"
                alt="Website icon. Click on it to go back to the Selection tab"
            />
            <h1>Custom {isFromExcel ? "Excel" : "Word"} Styling</h1>
        </div>
        <p>{lang("Change Word styles using the Office.JS API").replace("Word", isFromExcel ? "Excel" : "Word")}</p>
    </header>
    {#if showHomeDialog}
        <div
            class="dialog"
            in:fade={{ duration: 400, easing: cubicInOut }}
            out:fade={{ duration: 400, easing: cubicInOut }}
        >
            <div>
                <h2>{lang("Do you want to go back to the home?")}</h2>
                <p>{lang("You'll lose all the unsaved changes.")}</p>
                <div class="flex gap">
                    <button
                        onclick={async () => {
                            showHomeDialog = false;
                            // Since we'll re-render everything, we'll add a transition by creating a div that'll be on the top of the entire document, that'll scroll initially from the bottom to the top and then, after the re-render, from the top to the bottom.
                            const div = document.createElement("div");
                            div.classList.add("hideDocument");
                            document.body.append(div);
                            await new Promise(res => setTimeout(res, 15));
                            div.style.top = "0px";
                            await new Promise(res => setTimeout(res, 610));
                            forceReRender = `Entire-${Date.now()}`;
                            await new Promise(res => setTimeout(res, 100));
                            div.style.top = "100vh";
                            await new Promise(res => setTimeout(res, 610));
                            div.remove();
                        }}>{lang("Yes")}</button
                    >
                    <button
                        style="background-color: var(--input);"
                        onclick={() => (showHomeDialog = false)}
                        >{lang("No")}</button
                    >
                </div>
            </div>
        </div>
    {/if}
    {#if showLicenseDialog}
        <OpenSourceDialog callback={() => (showLicenseDialog = false)}
        ></OpenSourceDialog>
    {/if}
    {#if downloadLink}
        <div
            class="dialog"
            in:fade={{ duration: 400, easing: cubicInOut }}
            out:fade={{ duration: 400, easing: cubicInOut }}
        >
            <div>
                <h2>{lang("Download file")}</h2>
                <p>
                    {lang(
                        "Unfortunately, the platform you're using doesn't support file downloading. Copy the link displayed below, and then open it in your browser. Don't worry, your ultra-personal styles will always be private, and they aren't uploaded anywhere (all the download process happens locally).",
                    )}
                </p>
                <br />
                <div class="secondCard" style="overflow: scroll;">
                    <p style="white-space: pre;">{downloadLink}</p>
                </div>
                <br />
                <div class="flex gap">
                    <button
                        onclick={() =>
                            navigator.clipboard.writeText(
                                downloadLink as string,
                            )}>{lang("Copy")}</button
                    >
                    <button
                        style="background-color: var(--input);"
                        onclick={() => (downloadLink = false)}
                        >{lang("Close")}</button
                    >
                </div>
            </div>
        </div>
    {/if}

    {#if errorDialog}
        <div
            class="dialog"
            in:fade={{ duration: 400, easing: cubicInOut }}
            out:fade={{ duration: 400, easing: cubicInOut }}
        >
            <div>
                <h2>{lang("An error occurred")} :(</h2>
                <p>{errorDialog}</p>
                <br />
                <button onclick={() => (errorDialog = false)}
                    >{lang("Close")}</button
                >
            </div>
        </div>
    {/if}
    {#if showLicenseDialog}
        <OpenSourceDialog callback={() => (showLicenseDialog = false)}
        ></OpenSourceDialog>
    {/if}
    {#if isFromExcel}
        <Excel urlCallback={(url) => (downloadLink = url)}></Excel>
    {:else}
        <Word downloadLink={downloadLink}></Word>
    {/if}
    <br /><br />
    <div class="flex gap" style="flex-wrap: wrap;">
        <u onclick={() => changeTheme()}>{lang("Change theme")}</u>
        <u onclick={() => (showLicenseDialog = true)}
            >{lang("View open source licenses")}</u
        >
    </div>
    <p>
        {lang(
            "Word and Office are trademarks of Microsoft. This project is no way affiliated or endorsed by Microsoft",
        ).replace("Word", isFromExcel ? "Excel" : "Word")}.
    </p>
{/key}
