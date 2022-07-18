# Office Control IDs

The files in this repo provide the official Office IDs of Office controls and control groups that can be added to custom groups and ribbon tabs by Office Add-ins. Specifically, these IDs are used as the value of the `id` property in the [OfficeControl](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/group#control) and [OfficeGroup](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/customtab#officegroup) elements of the add-in's manifest.

For an overview of the feature, see [Integrate built-in Office buttons into custom control groups and tabs](https://docs.microsoft.com/office/dev/add-ins/design/built-in-button-integration).

> NOTE: There are limits on which controls, which Office applications, and which platforms support the `<OfficeControl>` and `<OfficeGroup>` elements. The files in this repo do not include IDs for unsupported scenarios.
>
> In particular, only PowerPoint add-ins are supported at this time.


## Instructions

To find the ID of a control or control group:

1. Open the folder for the Office application; for example, PowerPoint.
1. Open the Excel workbook for the platform; for example, `PowerPoint on the web Ribbon IDs.xlsx`.
1. To find a control ID:
    1. Open the **Controls** tab.
    1. Filter the **Ribbon Tab ID** column to include only the tab that has the built-in control that you want to insert in a custom group.
    1. Look in the **Office ID** column to find the ID. If it is not listed, then the control is not supported by the `<OfficeControl>` element and you cannot include it in your custom group.
1. To find a group ID:
    1. Open the **Single Line Ribbon Groups** tab.
    1. Filter the **Ribbon Tab ID** column to include only the tab that has the built-in group that you want to insert in a custom ribbon tab.
    1. Look in the **Office ID** column to find the ID *and make a note of it*, if it is there.
    1. Open the **Multiline Ribbon Groups** tab.
    1. Filter the **Ribbon Tab ID** column to include only the tab that has the built-in group that you want to insert in a custom ribbon tab.
    1. Look in the **Office ID** column to find the ID *and make a note of it*, if it is there. (If it is listed on both worksheets, it should be exactly the same ID.)
    1. If the group ID is on *both* worksheets, use it in the `<OfficeGroup>` element. It will be inserted in your custom ribbon tab on both the single and multiline ribbons.
    1. If the group ID is on *one* worksheet, then you can use it in the `<OfficeGroup>` element, but it will not appear in your custom ribbon tab on the other ribbon type.
    1. If the group ID is on *neither* worksheet, then the group is not supported by the `<OfficeGroup>` element and you cannot include it in your custom ribbon tab.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
