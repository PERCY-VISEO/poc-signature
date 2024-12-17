const fs = require('fs');

const path = require('path');
const JSZip = require('jszip');
const https = require('https');
const { parseString, Builder } = require('xml2js');

const IMAGE_URI = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";


const TYPE = 'type';
const HEIGHT = 'height';
const WIDTH = 'width';
const NAME = 'name';
const BUFFER = 'buffer';
const PATH = 'path';
const REL_ID = 'rel_id';

//const { DocxImager } = require('DocxImager');

class DocxImager {

    /**
     * Represents a DocxImager instance
     * @constructor
     */

    constructor() {
        this.zip = null;
    }


    replaceWithImageURL(image_uri, image_id, type, cbk) {
        this.__validateDocx();
        let req3 = https.request(image_uri, (res) => {
            let buffer = [];
            res.on('data', (d) => {
                buffer.push(d);
            });
            res.on('end', () => {
                fs.writeFileSync('t1.' + type, Buffer.concat(buffer));
                //res.headers['content-type']
                this.__replaceImage(Buffer.concat(buffer), image_id, type, cbk);
            });
        });

        req3.on('error', (e) => {
            console.error(e);
        });
        req3.end();
    }

    async __replaceImage(buffer, image_id, type, cbk) {
        //1. replace the image
        return new Promise((res, rej) => {
            try {
                let path = 'word/media/image' + image_id + '.' + type;
                this.zip.file(path, buffer);
                res(true);
            } catch (e) {
                rej();
            }
        });
    }

    /**
     * Load the DocxImager instance with the docx.
     * @param {String} path_or_buffer full path of the template docx, or the buffer
     * @returns {Promise}
     */
    async load(path_or_buffer) {
        return this.__loadDocx(path_or_buffer).catch((e) => {
            console.log(e);
        });
    }

    async __loadDocx(path_or_buffer) {
        let zip = new JSZip();
        let buffer = path_or_buffer;
        if (!Buffer.isBuffer(path_or_buffer)) {
            buffer = fs.readFileSync(path_or_buffer);
        }

        try {
            this.zip = await zip.loadAsync(buffer);
        } catch (e) {
            console.log('Error loading .docx file:', e);
            throw e;  // Rethrow the error for further handling
        }
    }


    static __cleanXML(xml) {

        //Simple variable match
        //({{|{(.*?){)(.*?)(}{1,2})(.*?)(?:[^}])(}{1})
        //1. ({{|{(.*?){)   - Match {{ or {<xmltgas...>{{
        //2. (.*?)          -   Match any character
        //3. (}}|}         -   Match } or }}
        //4. (.*?)          -   Match any character
        //5. (?:[^}])       -   KILLER: Stop matching
        //6. }               -   Include the Killer match
        let variable_regex = /({{|{(.*?){)(.*?)(}}|}(.*?)(?:[^}])})/g;
        let replacements = xml.match(variable_regex);

        // let replacements = xml.match(/({{#|{{#\/s)(?:(?!}}).)*|({{|{(.*?){)(?:(?!}}).)*|({{(.*?)#|{{#\/s)(?:(?!}}).)*/g);
        // let replacements = xml.match(/({{#|{{#\/s)(?:([^}}]).)*|({{|{(.*?){)(?:([^}}]).)*|({{(.*?)#|{{#\/s)(?:([^}}]).)*/g);
        // let replacements = xml.match(/({{#|{{#\/s)(?:(?!}}).)*|{{(?:(?!}}).)*|({{(.*?)#|{{(.*?)#\/s)(?:(?!}}).)*|{(.*?){(?:(?!}}).)*/g);//|({(.*?){(.*?)#|{{#\/s)(?:(?!}}).)*
        // let replacements = xml.match(/({(.*?){#|{(.*?){#\/s)(?:(?!}(.*?)}).)*|{(.*?){(?:(?!}(.*?)}).)*/g);
        let replaced_text;
        if (replacements) {

            for (let i = 0; i < replacements.length; i++) {
                if (replacements[i].includes("insert_image")) {
                    replaced_text = replacements[i].replace(/<\/w:t>.*?(<w:t>|<w:t [^>]*>)/g, '');
                    xml = xml.replace(replacements[i], replaced_text);
                }
            }
        }
        xml = xml.replace(/&quot;/g, '"');
        xml = xml.replace(/&gt;/g, '>');
        xml = xml.replace(/&lt;/g, '<');
        xml = xml.replace(/&amp;/g, '&');
        xml = xml.replace(/&apos;/g, '\'');

        return xml;
    }


    async __getVariableNames(context) {
        return new Promise(async (res, rej) => {
            try {
                // Step 1: Retrieve the document.xml content
                //console.log(context);
                let content = await this.zip.file('word/document.xml').async('nodebuffer');
                content = content.toString();
                content = DocxImager.__cleanXML(content);

                // Step 2: Initialize an empty object to store the variables
                let variables = {};
                // Step 3: Iterate over all keys in the context to extract images
                for (let imageName in context) {
                    //console.log(imageName);
                    // Dynamically create the regex to match the specific imageName in the document.xml
                    let regex = new RegExp(`{{insert_image\\s+${imageName}\\s+(\\S+)\\s+(\\d+)\\s+(\\d+)}}`, 'g');

                    // Step 4: Match the insert_image placeholders with the current image name
                    let matches = content.match(regex);

                    if (matches && matches.length) {
                        //console.log(matches);
                        // Extract variable names and attributes from each match
                        for (let i = 0; i < matches.length; i++) {
                            // Match the {{insert_image var_name type width height}} format
                            let tagMatch = matches[i].match(/{{insert_image\s+(\S+)\s+(\S+)\s+(\d+)\s+(\d+)}}/);

                            if (tagMatch) {
                                //console.log(tagMatch);
                                // Extract the variable name, type, width, and height
                                const variableName = tagMatch[1];
                                const type = tagMatch[2];
                                const width = tagMatch[3];
                                const height = tagMatch[4];

                                // Step 5: Build the variable entry for the final context
                                variables[variableName] = {
                                    [TYPE]: type,
                                    [WIDTH]: width,
                                    [HEIGHT]: height
                                };
                            } else {
                                console.warn(`Skipping malformed tag: ${matches[i]}`);
                            }
                        }
                    }
                }

                if (Object.keys(variables).length > 0) {
                    // Return the extracted variables
                    res(variables);
                } else {
                    rej(new Error('No insert_image placeholders found in document.xml'));
                }
            } catch (e) {
                console.error('Error in __getVariableNames:', e);
                rej(e);
            }
        });
    }



    async __getImageBuffer(path) {
        return new Promise((res, rej) => {
            try {
                const buffer = fs.readFileSync(path); // Read the file synchronously
                res(buffer); // Resolve with the image buffer
            } catch (e) {
                console.error(`Error reading local image at ${path}:`, e);
                rej(e); // Reject the promise with the error
            }
        });
    }

    async __getImages(variables, context) {
        return new Promise(async (res, rej) => {
            try {
                let image_map = {};
                for (let variable_name in variables) {
                    if (variables.hasOwnProperty(variable_name)) {
                        let path = context[variable_name];

                        // Initialize buffer variable for storing image data
                        let buffer;

                        if (path.startsWith('data:image/')) {
                            // Handle base64-encoded image
                            const base64Data = path.split(',')[1];  // Extract base64 data from the data URI
                            buffer = Buffer.from(base64Data, 'base64');
                        } 
                        else {
                            if (path.startsWith('http')) {
                                buffer = await this.__getImageBufferURL(path);
                            }
                            else {
                                // Handle local file path
                                const fs = require('fs').promises; // Use fs module for local file access
                                buffer = await fs.readFile(path);
                            }
                        }

                        // Generate a name for the image (e.g., image1.png or image1.jpg)
                        let type = variables[variable_name][TYPE];
                        let name = `image1.${type}`;
                        image_map[variable_name] = {
                            [NAME]: name,
                            [BUFFER]: buffer,
                            [TYPE]: type,
                        };
                    }
                }
                res(image_map);
            } catch (e) {
                console.error("Error in __getImages:", e);
                rej(e);
            }
        });
    }

    async __getImageBufferURL(path) {
        return new Promise((res, rej) => {
            try {
                let req = https.request(path, (result) => {
                    let buffer = [];
                    result.on('data', (d) => {
                        buffer.push(d);
                    });
                    result.on('end', () => {
                        res(Buffer.concat(buffer));
                    });
                });
                req.on('error', (e) => {
                    throw e;
                });
                req.end();
            } catch (e) {
                console.log(e);
                rej(e);
            }
        })
    }

    async _addContentType(final_context) {
        return new Promise(async (res, rej) => {
            try {
                // Read the [Content_Types].xml file as a buffer and convert it to a string
                let contentBuffer = await this.zip.file('[Content_Types].xml').async('nodebuffer');
                let content = contentBuffer.toString();

                // Ensure the <Types> tag exists in the content
                const typeTagRegex = /<Types.*?>/;
                const typeTagMatch = content.match(typeTagRegex);

                if (!typeTagMatch) {
                    throw new Error("Invalid [Content_Types].xml format: <Types> tag not found.");
                }

                // Initialize the string for new <Override> tags
                let overrideTags = '';

                // Loop over final_context and build the <Override> tags
                for (let var_name in final_context) {
                    if (final_context.hasOwnProperty(var_name)) {
                        const { name, type } = final_context[var_name];

                        // Validate NAME and TYPE, continue if missing
                        if (!name || !type) {
                            console.warn(`Skipping ${var_name}: Missing 'NAME' or 'TYPE' in final_context entry.`);
                            continue;
                        }

                        // Build <Override> tag for this entry
                        overrideTags += `<Override PartName="/word/media/${NAME}" ContentType="image/${TYPE}"/>`;
                    }
                }

                // If no <Override> tags were added, log and exit early
                if (!overrideTags) {
                    console.warn("No <Override> tags were added due to missing or invalid entries.");
                    return res(false);
                }

                // Insert the new <Override> tags into the [Content_Types].xml
                const updatedContent = content.replace(typeTagRegex, (match) => match + overrideTags);

                // Update the zip file with the modified content
                this.zip.file('[Content_Types].xml', updatedContent);

                // Successfully added the content type overrides
                res(true);
            } catch (e) {
                console.error("Error in _addContentType:", e);
                rej(e);
            }
        });
    }


    async _addImage(final_context) {
        return new Promise(async (res, rej) => {
            try {
                // Iterate over each entry in final_context to add the images
                for (let var_name in final_context) {
                    if (final_context.hasOwnProperty(var_name)) {
                        let o = final_context[var_name];

                        // Construct the image path where the image will be stored
                        let img_path = 'media/' + o[NAME];

                        // Add the image path to the entry object in final_context
                        o[PATH] = img_path;

                        // Add the image buffer to the 'word/media/' path inside the zip archive
                        this.zip.file('word/' + img_path, o[BUFFER]);
                    }
                }

                // Return success once all images have been added
                res(true);
            } catch (e) {
                // Log and reject in case of any errors during the process
                console.error("Error in _addImage:", e);
                rej(e);
            }
        });
    }


    async _addRelationship(final_context) {
        return new Promise(async (res, rej) => {
            try {
                // Read the 'document.xml.rels' file from the ZIP archive
                let content = await this.zip.file('word/_rels/document.xml.rels').async('nodebuffer');

                // Parse the XML content to work with it as an object
                parseString(content.toString(), (err, relation) => {
                    if (err) {
                        console.log("Error parsing document.xml.rels:", err);
                        rej(err);
                        return;
                    }

                    // Initialize the counter for relationship IDs
                    let cnt = relation.Relationships.Relationship.length;

                    // Loop over final_context to add image relationships
                    for (let var_name in final_context) {
                        if (final_context.hasOwnProperty(var_name)) {
                            let o = final_context[var_name];

                            // Construct a unique relationship ID
                            let rel_id = 'rId' + (++cnt);

                            // Store the relationship ID in the context object
                            o[REL_ID] = rel_id;

                            // Add the new relationship for the image (local image only)
                            relation.Relationships.Relationship.push({
                                $: {
                                    Id: rel_id,
                                    Type: IMAGE_URI,  // As per the spec for image relationships
                                    Target: o[PATH]   // Path to the image (which is local, as specified)
                                }
                            });
                        }
                    }

                    // Convert the modified object back into an XML string
                    let builder = new Builder();
                    let modifiedXML = builder.buildObject(relation);

                    // Update the 'document.xml.rels' file in the ZIP archive with the modified XML
                    this.zip.file('word/_rels/document.xml.rels', modifiedXML);

                    // Resolve the promise indicating success
                    res(true);
                });
            } catch (e) {
                // Log and reject in case of any error
                console.log("Error in _addRelationship:", e);
                rej(e);
            }
        });
    }


    static __getImgXMLElement(rId, height, width) {
        // width and height calculated assuming resolution as 96 dpi
        const calc_height = Math.round(9525 * height); // Calculate height in EMUs
        const calc_width = Math.round(9525 * width);   // Calculate width in EMUs
        const id = Math.floor(Math.random() * 10000) + 1;

        // Construct the XML element with proper namespaces
        return `<w:r>
                    <w:rPr>
                        <w:noProof/>
                    </w:rPr>
                    <w:drawing>
                        <wp:inline distT="0" distB="0" distL="0" distR="0">
                            <wp:extent cx="${calc_width}" cy="${calc_height}"/>
                            <wp:effectExtent l="0" t="0" r="0" b="0"/>
                            <wp:docPr id="${id}" name="Picture"/>
                            <wp:cNvGraphicFramePr>
                                <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                            </wp:cNvGraphicFramePr>
                            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                        <pic:nvPicPr>
                                            <pic:cNvPr id="1" name="Picture" descr=""/>
                                            <pic:cNvPicPr>
                                                <a:picLocks noChangeArrowheads="1"/>
                                            </pic:cNvPicPr>
                                        </pic:nvPicPr>
                                        <pic:blipFill>
                                            <a:blip r:embed="${rId}"/>
                                            <a:srcRect/>
                                            <a:stretch>
                                                <a:fillRect/>
                                            </a:stretch>
                                        </pic:blipFill>
                                        <pic:spPr bwMode="auto">
                                            <a:xfrm>
                                                <a:off x="0" y="0"/>
                                                <a:ext cx="${calc_width}" cy="${calc_height}"/>
                                            </a:xfrm>
                                            <a:prstGeom prst="rect">
                                                <a:avLst/>
                                            </a:prstGeom>
                                            <a:noFill/>
                                            <a:ln>
                                                <a:noFill/>
                                            </a:ln>
                                        </pic:spPr>
                                    </pic:pic>
                                </a:graphicData>
                            </a:graphic>
                        </wp:inline>
                    </w:drawing>
                </w:r>`;
    }


    async _addInDocumentXML(final_context) {
        return new Promise(async (res, rej) => {
            try {
                // Retrieve the 'document.xml' file content from the ZIP archive
                let content = await this.zip.file('word/document.xml').async('nodebuffer');
                content = content.toString();  // Convert buffer to string for manipulation

                // Clean the XML content using DocxImager's method (to ensure it's well-formed)
                content = DocxImager.__cleanXML(content);

                // Match all the <w:r> elements that contain the 'insert_image' keyword
                let matches = content.match(/(<w:r>.*?insert_image.*?<\/w:r>)/g);

                if (matches && matches.length > 0) {
                    // Create a log to track the number of replacements per image placeholder
                    let replacementCount = {};

                    // Loop over all matched image placeholders
                    for (let i = 0; i < matches.length; i++) {
                        let xml = matches[i];

                        // Use regex to extract the placeholder tag and its details (variable name, type, width, and height)
                        let regex = /{{insert_image\s+(\S+)\s+(\S+)\s+(\d+)\s+(\d+)}}/g;
                        let tag = regex.exec(xml);

                        while (tag) {
                            // Extract details from the tag (variable name, type, width, height)
                            let var_name = tag[1];
                            let type = tag[2]; // Type of the image (e.g., 'jpg', 'png')
                            let width = Number(tag[3]);
                            let height = Number(tag[4]);

                            // Fetch the corresponding image object from the final_context
                            let obj = final_context[var_name];

                            if (!obj) {
                                tag = regex.exec(xml);
                                continue; // Skip this iteration if the image data is missing
                            }

                            // Validate type and ensure it matches the context (optional validation step)
                            if (obj.type !== type) {
                                console.warn(`Type mismatch for variable "${var_name}": expected "${obj.type}", found "${type}"`);
                            }

                            // Replace the placeholder tag with the actual image XML element
                            xml = xml.replace(
                                tag[0],
                                `</w:t></w:r>${DocxImager.__getImgXMLElement(obj[REL_ID], height, width)}<w:r><w:t>`
                            );

                            // Count the number of replacements for this variable name (image)
                            if (replacementCount[var_name]) {
                                replacementCount[var_name]++;
                            } else {
                                replacementCount[var_name] = 1;
                            }

                            // Continue searching for more tags in the current XML segment
                            tag = regex.exec(xml);
                        }

                        // Replace the entire match with the modified XML
                        content = content.replace(matches[i], xml);
                    }

                    // After processing all matches, update the document.xml file in the ZIP archive
                    this.zip.file('word/document.xml', content);

                    // Log the number of replacements for each image placeholder
                    for (let var_name in replacementCount) {
                        if (replacementCount.hasOwnProperty(var_name)) {
                            console.log(`Placeholder "${var_name}" replaced ${replacementCount[var_name]} times by image "${final_context[var_name].name}"`);
                        }
                    }

                    res(true);
                } else {
                    // Reject the promise if no valid image placeholders are found
                    rej(new Error('Invalid Docx: No image placeholders found'));
                }
            } catch (e) {
                // Log and reject if there's an error
                console.log('Error in _addInDocumentXML:', e);
                rej(e);
            }
        });
    }


    async insertImage(context) {
        // a. get the list of all variables.

        let variables = await this.__getVariableNames(context);

        //b. download/retrieve images.
        let final_context = await this.__getImages(variables, context);

        // Deep merge image buffer, name, and meta data
        for (let var_name in final_context) {
            if (final_context.hasOwnProperty(var_name)) {
                // Merge image data (buffer and name)
                let imageData = final_context[var_name];

                // Ensure that all necessary meta data (like HEIGHT, WIDTH) is properly assigned
                if (variables[var_name]) {
                    // Assign only the meta data values that are present in the variables
                    imageData[HEIGHT] = variables[var_name][HEIGHT] || imageData[HEIGHT];
                    imageData[WIDTH] = variables[var_name][WIDTH] || imageData[WIDTH];
                }
            }
        }


        //1. insert entry in [Content-Type].xml
        await this._addContentType(final_context);

        //2. write image in media folder in word/
        /*let image_path = */await this._addImage(final_context);

        //3. insert entry in /word/_rels/document.xml.rels
        //<Relationship Id="rId3" Type=IMAGE_URI Target="media/image2.png"/>
        /*let rel_id = */await this._addRelationship(final_context);

        //4. insert in document.xml after calculating EMU.
        await this._addInDocumentXML(final_context);

        // http://polymathprogrammer.com/2009/10/22/english-metric-units-and-open-xml/
        // https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
    }
    /**
     * Saves the transformed docx.
     * @param {String} op_file_name Output file name with full path.
     * @returns {Promise}
     */
    async save(op_file_name) {
        if (!op_file_name) {
            op_file_name = './merged.docx';
        }

        return new Promise((res, rej) => {
            try {
                // Ensure that the document has the correct structure and integrity
                // Assuming `this.zip` contains the full document structure (including the correct files)
                this.zip.generateNodeStream({ streamFiles: true })
                    .pipe(fs.createWriteStream(op_file_name))
                    .on('finish', function () {
                        console.log('Document saved successfully at', op_file_name);
                        res();  // Resolve when the file has been written
                    })
                    .on('error', function (err) {
                        console.log('Error saving document:', err);
                        rej(err);  // Reject in case of error during the save process
                    });
            } catch (error) {
                console.log('Error in save function:', error);
                rej(error);  // Reject if there's an error in the function
            }
        });
    }

    __validateDocx() {
        if (!this.zip) {
            throw new Error('Invalid docx path or format. Please load docx.')
        }
    }

}




async function insertImageIntoDocx() {

    let docxImager = new DocxImager();
    const templatePath = path.resolve('./template.docx');
    const imagePath = "image.jpg";
    const outputPath = path.resolve('./output.docx');

    await docxImager.load(templatePath);
    //await docxImager.insertImage({ 'img1': imagePath });
    await docxImager.insertImage({ "img1": "https://www.cuteness.com/cuteness/17-super-wholesome-dog-memes-to-warm-your-heart/20e09723c51f41b5bebe1fbda21472c9.jpg" });
    await docxImager.insertImage({ 'img2': imagePath });
    await docxImager.save(outputPath);

}

insertImageIntoDocx().catch(err => {
    console.error('Error inserting image:', err);
});
