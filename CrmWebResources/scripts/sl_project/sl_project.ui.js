typeof (Softline) === 'undefined' && (window.Softline = {});
typeof (Softline.Subject) === 'undefined' && (Softline.Subject = {});

Softline.Subject.Buttons = {
    Translete: {
        command: fctx => {
            const isDirty = fctx.data.entity.getIsDirty();
            const id = fctx.data.entity.getId();
            if (!id || isDirty) {
                Xrm.Navigation.openAlertDialog('Save the card.', { height: 120, width: 120 })
                return;
            }
            const logicalname = fctx.data.entity.getEntityName();
            const pageInput = {
                pageType: "webresource",
                webresourceName: "sl_translete_form",
                data: `${id}&logicalname=${logicalname}`
            };
            const navigationOptions = {
                target: 2,
                height: { value: 80, unit: "%" },
                width: { value: 70, unit: "%" },
                position: 1,
                title: 'Translation Project'
            };
            Xrm.Navigation.navigateTo(pageInput, navigationOptions)
        },
        enable: fctx => {
            return true;
        }
    },
    UpdateFormatImages: {
        command: async (fctx) => {
            const isDirty = fctx.data.entity.getIsDirty();
            const reference = fctx.data.entity.getEntityReference();
            if (!reference.id || isDirty) {
                Xrm.Navigation.openAlertDialog('Save the card.', { height: 120, width: 120 })
                return;
            }
            Xrm.Utility.confirmDialog("Update format images ?", async () => await Softline.Subject.Buttons.UpdateFormatImages.updateFormat(fctx, reference));
        },

        updateFormat: async (fctx, reference) => {
            try {
                Xrm.Utility.showProgressIndicator("Loading");
                const updateFormatButton = Softline.Subject.Buttons.UpdateFormatImages;
                const request = updateFormatButton.retriveImageRequest(reference);
                const responce = await Xrm.WebApi.online.execute(request);
                const json = await responce.json();
                const answer = JSON.parse(json.responce);
                if (answer.IsError) {
                    Xrm.Utility.alertDialog(answer.Message);
                    return;
                }
                const images = answer.Images;
                if (!images || images.length === 0) {
                    Xrm.Utility.alertDialog("No images");
                    return;
                }
                const formatsName = images.map(i => i.BaseFolder);
                const allFormats = await updateFormatButton.getUploadFormat(reference.entityType, formatsName);
                const exclusivebit = fctx.getAttribute('sl_exclusivebit');
                const withWatermark = exclusivebit && !exclusivebit.getValue() ? true : false;
                const WATERMARK = "/WebResources/sl_watermark_big.png"
                const postData = [];

                const groupFormats = updateFormatButton.groupBy(allFormats, "sl_type@OData.Community.Display.V1.FormattedValue");
                for (let key in groupFormats) {
                    const formatsByName = groupFormats[key];
                    const imageResize = { BaseFolder: '', Resize: [] };
                    imageResize.BaseFolder = key;
                    let baseImages = images.filter(x => x.BaseFolder == key).map(x => x.Base).flat()
                    for (let i = 0; i < formatsByName.length; i++) {
                        const format = formatsByName[i];
                        const width = format.sl_width;
                        const height = format.sl_height;
                        const quality = format.sl_quality / 100;
                        const name = format.sl_name;
                        const id = format.sl_upload_formatid;
                        const withWatermarkFormat = format.sl_not_install_watermarkbit;
                        const resImagesPromise = [];
                        for (let i = 0; i < baseImages.length; i++) {
                            const baseImage = baseImages[i];
                            const promise = updateFormatButton.resizeImage(
                                baseImage.Base64,
                                width,
                                height,
                                quality,
                                baseImage.MimeType
                            )
                                .then((resImage) => {
                                    return withWatermark && !withWatermarkFormat
                                        ? updateFormatButton.setWatermark(resImage, WATERMARK)
                                        : resImage;
                                })
                                .then((imgWithWaterMark) => {
                                    return {
                                        NameForUrl: updateFormatButton.makeName(baseImage.FullName),
                                        FullName: baseImage.FullName,
                                        Base64: imgWithWaterMark,
                                        FolderName: `${name}`,
                                        MimeType: baseImage.MimeType,
                                        Formatid: id,
                                        Height: height,
                                        Width: width,
                                    };
                                });
                            resImagesPromise.push(promise);
                        }
                        const resizesImage64 = await Promise.all(resImagesPromise);
                        imageResize.Resize.push(...resizesImage64);
                    }
                    postData.push(imageResize);
                }
                let message = "";
                for (let data of postData) {
                    const uploadRequest = updateFormatButton.uploadImageRequest(reference, data);
                    const uploadResponce = await Xrm.WebApi.online.execute(uploadRequest);
                    const uploadJson = await uploadResponce.json();
                    const res = JSON.parse(uploadJson.responce);
                    const messageFormat = `format ${data.BaseFolder}`;
                    res.IsError
                        ? message += `${messageFormat} is error: ${res.Message}\n`
                        : message += `${messageFormat}: ${res.Message}\n`;
                }
                Xrm.Utility.alertDialog(message);
            } catch (e) {
                Xrm.Utility.alertDialog(e.message);
            }
            finally {
                Xrm.Utility.closeProgressIndicator();
            }
        },

        enable: (fctx) => {
            return true;
        },

        groupBy: (xs, key) => {
            return xs.reduce((rv, x) => {
                (rv[x[key]] = rv[x[key]] || []).push(x);
                return rv;
            }, {});
        },

        retriveImageRequest: (reference) => {
            return {
                regardingobject: {
                    "@odata.type": `Microsoft.Dynamics.CRM.${reference.entityType}`,
                    [`${reference.entityType}id`]: `${reference.id}`,
                },

                getMetadata: function () {
                    return {
                        boundParameter: null,
                        parameterTypes: {
                            regardingobject: {
                                typeName: `${reference.entityType}`,
                                structuralProperty: 5,
                            },
                        },
                        operationType: 0,
                        operationName: "sl_retrive_images",
                    };
                },
            };



        },

        uploadImageRequest: (reference, postData) => {
            return {
                regardingobject: {
                    "@odata.type": `Microsoft.Dynamics.CRM.${reference.entityType}`,
                    [`${reference.entityType}id`]: `${reference.id}`,
                },
                images: JSON.stringify(postData),
                getMetadata: function () {
                    return {
                        boundParameter: null,
                        parameterTypes: {
                            images: {
                                typeName: "Edm.String",
                                structuralProperty: 1,
                            },
                            regardingobject: {
                                typeName: `${reference.entityType}`,
                                structuralProperty: 5,
                            },
                        },
                        operationType: 0,
                        operationName: "sl_upload_image",
                    };
                },
            };
        },

        getUploadFormat: async (entityType, formats) => {
            const responce = await Xrm.Utility.getEntityMetadata("sl_upload_format", [
                "sl_entity",
            ]);

            const optionSetValue = Object.entries(
                responce.Attributes._collection["sl_entity"].OptionSet
            ).find((x) => x[1].text === entityType);

            const entity = optionSetValue.length > 0 && optionSetValue[0];
            const map = { Images: 102690000, Plans: 102690001 }
            const typeArr = formats.map(f => map[f]).filter(f => f);
            const formatFilter = typeArr.map((key) => `<value>${key}</value>`).join("");
            const query = `<fetch no-lock='true' >
                                            <entity name='sl_upload_format' >
                                              <attribute name='sl_height' />
                                              <attribute name='sl_name' />
                                              <attribute name='sl_width' />
                                              <attribute name='sl_quality' />
                                              <attribute name='sl_type' />
                                              <attribute name='sl_not_install_watermarkbit' />
                                              <filter type='and' >
                                                <condition attribute='statecode' operator='eq' value='0' />
                                                <condition attribute='sl_entity' operator='eq' value='${entity}' />
                                                <condition attribute='sl_type' operator='in' >
                                                  ${formatFilter}
                                                </condition>
                                              </filter>
                                            </entity>
                                          </fetch>`;
            const responceFormat = await Xrm.WebApi.retrieveMultipleRecords(
                "sl_upload_format",
                `?fetchXml=${query}`
            );
            return responceFormat.entities;
        },

        resizeImage: async (file, width, height, quality, mimeType) => {
            const image = new Image();
            image.src = file;
            await new Promise((resolve) => {
                image.onload = resolve;
            });

            const canvas = document.createElement("canvas"),
                ctx = canvas.getContext("2d"),
                oc = document.createElement("canvas"),
                octx = oc.getContext("2d");
            const inputWidth = image.width;
            const inputHeight = image.height;
            const inputCoef = inputWidth / inputHeight;
            const coef = width / height;
            const step = 0.5;
            if (inputCoef < coef) {
                canvas.width = width;
                canvas.height = (width * inputHeight) / inputWidth;
                if (inputWidth * step > width) {
                    const mul = 1 / step;
                    let cur = {
                        width: Math.floor(inputWidth * step),
                        height: Math.floor(inputHeight * step),
                    };

                    oc.width = cur.width;
                    oc.height = cur.height;

                    octx.drawImage(image, 0, 0, cur.width, cur.height);

                    while (cur.width * step > width) {
                        cur = {
                            width: Math.floor(cur.width * step),
                            height: Math.floor(cur.height * step),
                        };
                        octx.drawImage(
                            oc,
                            0,
                            0,
                            cur.width * mul,
                            cur.height * mul,
                            0,
                            0,
                            cur.width,
                            cur.height
                        );
                    }
                    ctx.drawImage(oc, 0, 0, cur.width, cur.height, 0, 0, canvas.width, canvas.height);
                } else {
                    ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
                }

                image.src = canvas.toDataURL(mimeType, quality);
                await new Promise((resolve) => {
                    image.onload = resolve;
                });

                const center = (image.height - height) * 0.5;
                // const outputY = center < 0 ? 0 : center;
                canvas.height = height;

                ctx.drawImage(image, 0, center, width, height, 0, 0, width, height);

                return canvas.toDataURL(mimeType, quality);
            } else {
                canvas.width = (height * inputWidth) / inputHeight;
                canvas.height = height;
                if (inputHeight * step > height) {
                    const mul = 1 / step;
                    let cur = {
                        width: Math.floor(inputWidth * step),
                        height: Math.floor(inputHeight * step),
                    };

                    oc.width = cur.width;
                    oc.height = cur.height;

                    octx.drawImage(image, 0, 0, cur.width, cur.height);

                    while (cur.height * step > height) {
                        cur = {
                            width: Math.floor(cur.width * step),
                            height: Math.floor(cur.height * step),
                        };
                        octx.drawImage(
                            oc,
                            0,
                            0,
                            cur.width * mul,
                            cur.height * mul,
                            0,
                            0,
                            cur.width,
                            cur.height
                        );
                    }
                    ctx.drawImage(oc, 0, 0, cur.width, cur.height, 0, 0, canvas.width, canvas.height);
                } else {
                    ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
                }

                image.src = canvas.toDataURL(mimeType, quality);

                await new Promise((resolve) => {
                    image.onload = resolve;
                });

                const center = (image.width - width) * 0.5;
                canvas.width = width;
                ctx.drawImage(image, center, 0, width, height, 0, 0, width, height);
                return canvas.toDataURL(mimeType, quality);
            }
        },

        setWatermark: async (baseImage, watermarkImage) => {
            const request = await fetch(baseImage);
            const blob = await request.blob();
            const img = await watermark([blob, watermarkImage], { type: blob.type }).image(
                watermark.image.lowerRight()
            );
            return img.src;
        },

        makeName: (oldName) => {
            const match = oldName.match(".jpg|.jpeg");
            const length = match.index;
            let result = "";
            const characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            const charactersLength = characters.length;
            for (let i = 0; i < length; i++) {
                result += characters.charAt(Math.floor(Math.random() * charactersLength));
            }
            return result + oldName.substring(length);
        }
    },
    SaveCoordinates: {
        command: fctx => {
            Xrm.Utility.confirmDialog("Save new coordinates ?", () => Softline.Subject.Buttons.SaveCoordinates.saveCoordinates(fctx))
        },
        enable: fctx => {
            return true;
        },
        saveCoordinates: async fctx => {
            const mapControl = fctx.getControl("WebResource_map");
            if (!mapControl) return;
            const contentWindow = await mapControl.getContentWindow();
            const markers = contentWindow.markers;
            const marker = markers.length > 0 ? markers[0] : null;
            if (!marker) return;
            const latAttr = fctx.getAttribute("sl_google_map_lat");
            const lngAttr = fctx.getAttribute("sl_google_map_lng");
            latAttr && latAttr.setValue(marker.position.lat());
            lngAttr && lngAttr.setValue(marker.position.lng());
            fctx.data.entity.save();
            Xrm.Utility.alertDialog("Coordinates saved");
        },
    },
    PrintForm: {
        command: async (fctx, ids) => {
            debugger;
            //Xrm.Utility.showProgressIndicator("Loading");
            //const projectid = ids[0];
            //const request = Softline.Subject.Buttons.PrintForm.printFormRequest(projectid, 1);
            //const responce = await Xrm.WebApi.online.execute(request);
            //const json = await responce.json();
            //const data = JSON.parse(json.responce);
            //if (data.IsError) {
            //    Xrm.Utility.closeProgressIndicator();
            //    Xrm.Utility.alertDialog(data.Message);
            //    return;
            //}
            //const byteArray = Softline.Subject.Buttons.PrintForm.base64ToArrayBuffer(data.File); 
            //Xrm.Utility.closeProgressIndicator();          
            //Softline.Subject.Buttons.PrintForm.download(byteArray, data.FileName, "application/msword");
            const json = `{
    "objectName": "sky",
	"preview": "https://static.tildacdn.com/tild3830-3562-4865-b562-643637396531/view01_final.jpg",
	"detailsLabel": "Details",
	"details": {"City": "Limassol", "Area": "Potamos Germasogeias", "Completion": "Under construction", "Distance to sea": "300 m", "Prices from": "800,000 �", "Completion date": "Jan 2021"},
	"majorBenefitsLabel": "Major benefits",
	"majorBenefits": "<ul style='padding-left: 0;'><li>landmark architecture from the leading London bureau;</li> <li>the only high-rise condominium of Limassol with no noise from the coastal road;</li><li>breathtaking unobstructed sea views;</li><li>hotel-style facilities: concierge, SPA, outdoor pool, heated indoor pool, gym, tennis court, green area, and underground parking;</li></ul>",
	"websiteUrl": "bbf.com",
	"phoneLabel": "T",
	"phone": "+357 25 315 300",
	"faxLabel": "F",
	"fax": "+357 25 315 301",
	"emailLabel": "E",
	"email": "info@bbf.com",
	"address": "28 Ambelakion Street, Germasogeia 4046 Limassol, Cyprus P.O. Box 70649, 3801 Limassol",
	"tagline": "Build. Better. Future.",
	"metadata": {
		"withLogo": false,
		"mainTableColumns": [
			{"name":"�","unit":false},
			{"name":"IBP","unit":false},
			{"name":"Floor","unit":false},
			{"name":"Type","unit":false},
			{"name":"Bedrooms","unit":false},
			{"name":"Indoor area","unit":"sq. m."},
			{"name":"Cov. veranda","unit":"sq. m."},
			{"name":"Uncov. veranda","unit":"sq. m."},
			{"name":"Roof terrace","unit":"sq. m."},
			{"name":"Common area","unit":"sq. m."},
			{"name":"Total area","unit":"sq. m."},
			{"name":"Pool","unit":false},
			{"name":"PRICE","unit":"�"},
			{"name":"VAT","unit":false}
		]		
	},
    "mainTable": [{
		"blockName": "Block A",
		"flats": [
					[
						[ "� 101","","1","Appartment","2","75.6","18.7","0.0","0.0","10.5","104.7","Common","Sold","Yes"],
						[ "� 101","","1","Appartment","2","75.6","18.7","0.0","0.0","10.5","104.7","Common","Sold","Yes"],
						[ "� 101","","1","Appartment","2","75.6","18.7","0.0","0.0","10.5","104.7","Common","Sold","Yes"]
					],
					[
						[ "� 101","","1","Appartment","2","75.6","18.7","0.0","0.0","10.5","104.7","Common","Sold","Yes"],
						[ "� 101","","1","Appartment","2","75.6","18.7","0.0","0.0","10.5","104.7","Common","Sold","Yes"],
						[ "� 101","","1","Appartment","2","75.6","18.7","0.0","0.0","10.5","104.7","Common","Sold","Yes"]
					]
				]
	},
	{
		"blockName": "Block B",
		"flats": [
					[
						[ "� 101","","1","Appartment","2","75.6","18.7","0.0","0.0","10.5","104.7","Common","Sold","Yes"],
						[ "� 101","","1","Appartment","2","75.6","18.7","0.0","0.0","10.5","104.7","Common","Sold","Yes"],
						[ "� 101","","1","Appartment","2","75.6","18.7","0.0","0.0","10.5","104.7","Common","Sold","Yes"]
					],
					[
						[ "� 101","","1","Appartment","2","75.6","18.7","0.0","0.0","10.5","104.7","Common","Sold","Yes"],
						[ "� 101","","1","Appartment","2","75.6","18.7","0.0","0.0","10.5","104.7","Common","Sold","Yes"],
						[ "� 101","","1","Appartment","2","75.6","18.7","0.0","0.0","10.5","104.7","Common","Sold","Yes"]
					]
				]
	}]
}`
            sessionStorage.setItem("dataPrintForm", JSON.stringify(json));
            const pageInput = {
                pageType: "webresource",
                webresourceName: "sl_/reports/Project.html",
            };
            const navigationOptions = {
                target: 2,
                height: { value: 80, unit: "%" },
                width: { value: 70, unit: "%" },
                position: 1,
                title: 'Project'
            };
            Xrm.Navigation.navigateTo(pageInput, navigationOptions)
        },
        enable: fctx => {
            return true;
        },

        download: (data, filename, type) => {
            var file = new Blob([data], { type: type });
            if (window.navigator.msSaveOrOpenBlob)
                window.navigator.msSaveOrOpenBlob(file, filename);
            else {
                var a = document.createElement("a"),
                    url = URL.createObjectURL(file);
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                setTimeout(function () {
                    document.body.removeChild(a);
                    window.URL.revokeObjectURL(url);
                }, 0);
            }
        },

        base64ToArrayBuffer: (base64) => {
            const binaryString = window.atob(base64);
            const len = binaryString.length;
            const bytes = new Uint8Array(len);
            for (let i = 0; i < len; i++) {
                bytes[i] = binaryString.charCodeAt(i);
            }
            return bytes.buffer;
        },

        printFormRequest: (projectid, templateid) => {
            return {
                data: {
                    "@odata.type": `Microsoft.Dynamics.CRM.sl_project`,
                    [`sl_projectid`]: projectid,
                },
                templateid: templateid,
                getMetadata: function () {
                    return {
                        boundParameter: null,
                        parameterTypes: {
                            templateid: {
                                typeName: "Edm.Int",
                                structuralProperty: 1,
                            },
                            data: {
                                typeName: `sl_project`,
                                structuralProperty: 5,
                            },
                        },
                        operationType: 0,
                        operationName: "sl_get_print_form",
                    };
                },
            };
        }
    }
}