typeof (Softline) === 'undefined' && (window.Softline = {});
typeof (Softline.SpDocuments) === 'undefined' && (Softline.SpDocuments = {});

Softline.SpDocuments.Buttons = {
    UploadImage: {
        command: fctx => {
            const isDirty = fctx.data.entity.getIsDirty();
            const id = fctx.data.entity.getId();
            if (isDirty || !id) {
                Xrm.Navigation.openAlertDialog('Save the card.', { height: 120, width: 120 })
                return;
            }
            const logicalname = fctx.data.entity.getEntityName();
            const exclusivebit = fctx.getAttribute('sl_exclusivebit');
            const withWatermark = exclusivebit && !exclusivebit.getValue() ? true : false;
            const pageInput = {
                pageType: "webresource",
                webresourceName: "sl_imageLoader",
                data: `${id}&logicalname=${logicalname}&withWatermark=${withWatermark}`
            };
            const navigationOptions = {
                target: 2,
                height: 318,
                width: 598,
                position: 1,
                title: 'Upload Images'
            };
            const successCallback = () => fctx.data.refresh()
            Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(successCallback);
        },
        enable: fctx => {
            const logicalname = fctx.data.entity.getEntityName();
            return ['sl_unit', "sl_project"].includes(logicalname);
        }
    }
}