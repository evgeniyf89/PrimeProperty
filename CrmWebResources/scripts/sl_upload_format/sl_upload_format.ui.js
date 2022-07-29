typeof (Softline) === 'undefined' && (window.Softline = {});
typeof (Softline.UploadFormat) === 'undefined' && (Softline.UploadFormat = {});

Softline.UploadFormat.Buttons = {
    UpdateImages: {
        command: fctx => {
            debugger;
            const isDirty = fctx.data.entity.getIsDirty();
            const reference = fctx.data.entity.getEntityReference();
            if (!reference.id || isDirty) {
                Xrm.Navigation.openAlertDialog('Save the card.', { height: 120, width: 120 })
                return;
            }

            Xrm.Utility.confirmDialog("Update all images ?",
                async () => {
                    try {
                        Xrm.Utility.showProgressIndicator("Loading");
                        const updateImages = Softline.UploadFormat.Buttons.UpdateImages;
                        const format = updateImages.getUploadFormat(fctx);
                        const objectsWithPicture = await updateImages.retriveAllObjects(format);
                        let updateCount = 0;                       
                        const count = objectsWithPicture.length;
                        for (let i = 0; i < count; i++) {
                            const object = objectsWithPicture[i];
                            object.id = object[`${format.sl_entity}id`];
                            object.entityType = format.sl_entity;
                            const resFormat = await updateImages.updateFormat(object, format);
                            if (!resFormat.IsError) {
                                ++updateCount;
                            } else {
                                const delay = resFormat.RetryAfter;
                                delay && await updateImages.sleep(delay * 1000);
                                await updateImages.updateFormat(object, format);
                            }
                            Xrm.Utility.showProgressIndicator(`${updateCount} update from ${count}`);
                        }
                        const messageForProject = `Update: ${updateCount}, errors: ${count - updateCount}`;
                        Xrm.Utility.alertDialog(messageForProject);
                    } catch (e) {
                        Xrm.Utility.alertDialog(e.message);
                    }
                    finally {
                        Xrm.Utility.closeProgressIndicator();
                    }
                });
        },
        sleep: (ms) => {
            return new Promise(resolve => setTimeout(resolve, ms));
        },
        retriveAllObjects: async (format) => {
            const from = format.sl_entity == "sl_unit"
                ? "sl_propertyid"
                : `${format.sl_entity}id`;
            let page = 1;
            let allObjects = [];
            var now = new Date();
            now.setDate(now.getDate() - 5);         
            while (true) {
                const query =
                    `<fetch distinct='true' no-lock='true' page='${page} ' >
                          <entity name='${format.sl_entity}' >
                          <attribute name='${format.sl_entity}id' />                   
                            <link-entity name='sl_picture' from='${from}' to='${format.sl_entity}id' link-type='inner' alias='aa'>
                              <filter>
                                <condition attribute='createdon' operator='on-or-before' value='${now.format("MM-dd-yyyy")}' />
                                <condition attribute='sl_upload_formatid' operator='eq' value='${format.sl_upload_formatid}' />
                              </filter>
                            </link-entity>
                          </entity>
                    </fetch>`;
                const response = await Xrm.WebApi.retrieveMultipleRecords(
                    `${format.sl_entity}`,
                    `?fetchXml=${query}`
                );
                allObjects.push(...response.entities);
                page = page + 1;
                cookie = response.fetchXmlPagingCookie;
                if (response.entities.length < 5000) break;
            }
            return allObjects;
        },

        updateFormat: async (entity, format) => {
            const updateImages = Softline.UploadFormat.Buttons.UpdateImages;
            const request = updateImages.retriveImageRequest(entity, format);
            const responce = await Xrm.WebApi.online.execute(request);
            const json = await responce.json();
            const answer = JSON.parse(json.responce);
            if (answer.IsError) {
                console.error(answer.Message);
                return answer;
            }
            const images = answer.Images;
            if (!images || images.length === 0) {
                return answer
            }
            const withWatermark = entity.sl_exclusivebit && !entity.sl_exclusivebit ? true : false;
            const WATERMARK = "/WebResources/sl_watermark_big.png";
            const imageResize = { BaseFolder: format.sl_type_name, Resize: [] };

            const width = format.sl_width;
            const height = format.sl_height;
            const quality = format.sl_quality / 100;
            const name = format.sl_name;
            const id = format.sl_upload_formatid;
            const withWatermarkFormat = format.sl_not_install_watermarkbit;
            const resImagesPromise = [];
            const baseImages = images.map(x => x.Base).flat();
            for (let i = 0; i < baseImages.length; i++) {
                const baseImage = baseImages[i];
                const promise = updateImages.resizeImage(
                    baseImage.Base64,
                    width,
                    height,
                    quality,
                    baseImage.MimeType
                )
                    .then((resImage) => {
                        return withWatermark && !withWatermarkFormat
                            ? updateImages.setWatermark(resImage, WATERMARK)
                            : resImage;
                    })
                    .then((imgWithWaterMark) => {
                        return {
                            NameForUrl: updateImages.makeName(baseImage.FullName),
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

            const uploadRequest = updateImages.uploadImageRequest(entity, imageResize);
            const uploadResponce = await Xrm.WebApi.online.execute(uploadRequest);
            const uploadJson = await uploadResponce.json();
            const res = JSON.parse(uploadJson.responce);
            if (res.IsError) {
                console.error(res.Message);
            }
            return res;
        },

        retriveImageRequest: (reference, formatImage) => {
            return {
                regardingobject: {
                    "@odata.type": `Microsoft.Dynamics.CRM.${reference.entityType}`,
                    [`${reference.entityType}id`]: `${reference.id}`,
                },
                format: {
                    "@odata.type": `Microsoft.Dynamics.CRM.${formatImage.entityType}`,
                    [`${formatImage.entityType}id`]: `${formatImage.sl_upload_formatid}`,
                    "sl_name": formatImage.sl_name,
                    "sl_type": formatImage.sl_type
                },
                getMetadata: function () {
                    return {
                        boundParameter: null,
                        parameterTypes: {
                            regardingobject: {
                                typeName: `${reference.entityType}`,
                                structuralProperty: 5,
                            },
                            format: {
                                typeName: `${formatImage.entityType}`,
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

        getUploadFormat: (fctx) => {
            const getValue = (attrName) => {
                const attr = fctx.getAttribute(attrName);
                return attr && attr.getValue();
            }
            const reference = fctx.data.entity.getEntityReference();
            const formatData = ["sl_width", "sl_height", "sl_quality", "sl_name", "sl_entity", "sl_type"];
            const format = {};
            for (let data of formatData) {
                const value = getValue(data);
                if (!value) {
                    Xrm.Utility.alertDialog(`${data} is empty`);
                    return;
                }
                format[data] = value;
            }
            const entityCode = format.sl_entity;

            format.sl_upload_formatid = reference.id;
            format.entityType = reference.entityType;
            format.sl_not_install_watermarkbit = getValue("sl_not_install_watermarkbit");
            format.sl_entity = entityCode == 102690001
                ? "sl_project"
                : entityCode == 102690000 ? "sl_unit" : '';
            format.sl_type_name = format.sl_type == 102690000
                ? "Images"
                : format.sl_type == 102690001 ? "Plans" : '';
            return format;
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
        },

        enable: fctx => {
            const entity = fctx.getAttribute("sl_entity");
            const value = entity && entity.getValue();
            return [102690000, 102690001].includes(value);
        }
    }
}