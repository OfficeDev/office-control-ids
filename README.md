# Office Control IDs

The files in this repo provide the official Office IDs of Office controls and control groups that can be added to custom groups and ribbon tabs by Office Add-ins. Specifically, these IDs are used as the value of the `id` property in the [OfficeControl](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/group#control) and [OfficeGroup](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/customtab#officegroup) elements of the add-in's manifest.

For an overview of the feature, see [Integrate built-in Office buttons into custom control groups and tabs](https://docs.microsoft.com/office/dev/add-ins/design/built-in-button-integration).

> NOTE: There are limits on which controls, which Office applications, and which platforms support the `<OfficeControl>` and `<OfficeGroup>` elements. The files in this repo do not include IDs for unsupported scenarios.

## Instructions

To find the ID of a control or control group:

1. Open the folder for the Office application; for example, PowerPoint.
1. Open the Excel workbook for the platform; for example, `PowerPoint on the web Ribbon IDs.xlsx`.
1. To find a control ID:

    a. Open the **Controls** tab.
    b. Filter the **Ribbon Tab ID** column to include only the tab that has the built-in control that you want to insert in a custom group.
    c. Look in the **Office ID** column to find the ID. If it is not listed, then the control is not supported by the `<OfficeControl>` element and you cannot include it in your custom group.

1. To find a group ID:

    a. Open the **Single Line Ribbon Groups** tab.
    b. Filter the **Ribbon Tab ID** column to include only the tab that has the built-in group that you want to insert in a custom ribbon tab.
    c. Look in the **Office ID** column to find the ID *and make a note of it*, if it is there.
    d. Open the **Multiline Ribbon Groups** tab.
    e. Filter the **Ribbon Tab ID** column to include only the tab that has the built-in group that you want to insert in a custom ribbon tab.
    f. Look in the **Office ID** column to find the ID *and make a note of it*, if it is there. (If it is listed on both worksheets, it should be exactly the same ID.)
    g. If the group ID is on *both* worksheets, use it in the `<OfficeGroup>` element. It will be inserted in your custom ribbon tab on both the single and multiline ribbons.
    h. If the group ID is on *one* worksheet, then you can use it in the `<OfficeGroup>` element, but it will not appear in your custom ribbon tab on the other ribbon type.
    i. If the group ID is on *neither* worksheet, then the group is not supported by the `<OfficeGroup>` element and you cannot include it in your custom ribbon tab.

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
