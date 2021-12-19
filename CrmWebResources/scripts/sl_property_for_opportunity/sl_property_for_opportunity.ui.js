typeof Softline === "undefined" && (window.Softline = {});
typeof Softline.PropertyForOpportunity && (Softline.PropertyForOpportunity = {});
Softline.PropertyForOpportunity.Buttons = {
    Calendar: {
        command: (fctx) => {
            const isDirty = fctx.data.entity.getIsDirty();
            const id = fctx.data.entity.getId();
            if (!id || isDirty) {
                Xrm.Navigation.openAlertDialog("Save the card.", { height: 120, width: 120 });
                return;
            }
            const listingid = fctx.getAttribute("sl_listingid")?.getValue()[0].id ?? '';
            const propertyRef = fctx.getAttribute("sl_propertyid")?.getValue()[0];
            const opportunityid = fctx.getAttribute("sl_opportunityid")?.getValue()[0].id ?? '';
            const start = fctx.getAttribute("sl_date_from").getValue()?.format("yyyy.MM.dd") ?? '';
            const end = fctx.getAttribute("sl_date_to").getValue()?.format("yyyy.MM.dd") ?? '';           

            const pageInput = {
                pageType: "webresource",
                webresourceName: "sl_calendar.html",
                data: `${propertyRef.id}&logicalname=${propertyRef.entityType}&opportunityid=${opportunityid}&listingid=${listingid}&start=${start}&end=${end}&propertyOppid=${id}`,
            };
            const navigationOptions = {
                target: 2,
                height: 550,
                width: 850,             
                position: 1,
                title: "Calendar",
            };
            const successCallback = () => fctx.data.refresh()
            Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(successCallback);
        },
        enable: async (fctx) => {
            const opportunityAttr = fctx.getAttribute("sl_opportunityid");
            const opp = opportunityAttr && opportunityAttr.getValue();
            if (!opp) return false;
            const responce = await Xrm.WebApi.retrieveRecord(opp[0].entityType, opp[0].id, "?$select=sl_promotion_typecode");
            return responce.sl_promotion_typecode == 102690002;
        },
    }
};

