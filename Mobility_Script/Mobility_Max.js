// ==UserScript==
// @name         Mobility-Maximized
// @namespace    http://tampermonkey.net/
// @version      1.0.35
// @description  Enhance part handling and UI in Mobility.
// @author       moorpatx@ 1/12/25
// @icon         https://drive-render.corp.amazon.com/view/pjbyrne@/Script%20Logos/Rocket.png
// @match        https://mobility.amazon.com/part/search*
// @match        https://prod.us-east-1.mobility.scm.aws.dev/part/search*
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.4/dist/xlsx.full.min.js
// @exclude      https://mobility.amazon.com/part/part/*
// @exclude      https://mobility.amazon.com/part/bin/*
// @downloadURL  https://drive.corp.amazon.com/documents/moorpatx@/Mobility-Maximized-1.0.35.user.js#bypass=true
// @updateURL    https://drive.corp.amazon.com/documents/moorpatx@/Mobility-Maximized-1.0.35.user.js#bypass=true
// @grant        GM_addStyle
// @grant        GM_info
// @grant        GM_addStyle
// @grant        GM_getResourceText
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_xmlhttpRequest
// ==/UserScript==


// CHANGELOG
// 1.0.1  - initial release
// 1.0.2  - Minor style changes in buttons and addition of Build ID to part summary.
// 1.0.3  - Addition of San Requirement Color-coding summary which appears on the NavBar at top of Mobility Parts page + fetching of scanned serial Model in the Scan Multiple box
// 1.0.4  - Added slide up notification tab for San Requirement status function once all have loaded
// 1.0.5  - Pop-up prompt for state change and Dark Mode added.
// 1.0.6  - Dark Mode preference state saved on reload. Popup Prompt added for TOA/Local bin transfer functions.
// 1.0.7  - Added side-scroll buttons and amended the TOA/Local transfer prompt logic.
// 1.0.8  - MXP Bins added to TOA/LOCAL Bin dropdopwn. Popup prompt logic changed for TOA\LOCAL Bins.
// 1.0.9  - Removed csv download for "General" Transfer option.
// 1.0.10  - Can now click serial cells in Last Users table to get the full audit history in a pop-up window, Bulk Model Change dropdown option added to toolbar, column display preferences button added.
// 1.0.11 - Copy Bin/Model/MPN/APN/PO text by clicking relevant cell.
// 1.0.12 - Copy SHIPMENT ID, BUILD ID, RMA ID, OUTBOUND RMA ID, TRACKING, OUTBOUND TRACKING ID.
// 1.0.13 - Added PENDING_RMA, PENDING_SAN options to state change dropdown.
// 1.0.14 - Added Tote ID Header/Hyperlinks in side menu for easy fitering/opening in new tab.
// 1.0.15 - Some fixes on the Tote ID function
// 1.0.16 - Adding matching/highlight function to serial scan box for matching serials on current page + fixing/enhancing row size function (Upon select of option, it auto-reloads the page with that number)
// 1.0.17 - Merged Slide Left/Up buttons into 1 diagonal button that hides both elements at same time, added slider elements instead of checkboxes to column display pref.
// 1.0.18 - Fix in the Description Shortener logic
// 1.0.19 - Fix in Scan Multiple Serials model display logic - added popup with relevant model.
// 1.0.20 - PO Number Filtering in side menu added
// 1.0.21 - Advanced Edit issue fixed, Bin Lookup Window added (Click "Site" cell to trigger)
// 1.0.22 - Bin Lookup URL syntax fixed
// 1.0.23 - Download Link fixed - shortened file name
// 1.0.24 - Download Link fixed - shortened file name + config, tweaked Bin Lookup logic
// 1.0.25 - Fixes + configs: Download button logic, scan multiple serials "-" added as an allowed character.
// 1.0.26 - LHR Local Bins added
// 1.0.27 - Copy function added to description - now only bulk expand/collapse function for desc
// 1.0.28 - Part summary data fix + Model Description shortener.
// 1.0.29 - Copy cell data logic fixed
// 1.0.30 - Copy cell data logic fixed
// 1.0.31 - PHX Sites added to bin lookup
// 1.0.32 - Compatability with Final Destination script.
// 1.0.33 - Scanned serial stays at top.
// 1.0.34 - Model Desc copy fix.
// 1.0.35 - Bin change syntax fix.
// 1.0.36 - Added user preferences storage and export/import functionality [working progress] *
// 1.0.37 - Added keyboard shortcuts for common actions [working progress] *
// 1.0.38 - Added enhanced search filters and analytics dashboard [working progress] *
// 1.0.39 - Added batch operations history and improved error handling [working progress] *
// 1.0.40 - Added auto-refresh and custom layouts functionality [working progress] *



/////////🏁 SCRIPT START ///////////////////////



(function () {






    /////🚀 EDIT MOBILITY TITLE TO SHOW SCRIPT NAME /////////////////


    function addMaximizedText() {
        // Get the element to modify its text content and style
        const brandElement = document.querySelector('.brand.mobility-brand');

        // Modify the text content
        brandElement.textContent += ' '; // Add a space before appending the text
        const maximizedSpan = document.createElement('span');
        maximizedSpan.textContent = 'Maximized';
        maximizedSpan.style.fontSize = '14px';
        maximizedSpan.style.fontFamily = 'Sacramento, cursive'; // Lobster font family (handwriting style)
        maximizedSpan.style.letterSpacing = '0.5px'; // Add some letter spacing for a retro effect
        brandElement.appendChild(maximizedSpan);

        // Create an image element
        const imageElement = document.createElement('img');
        imageElement.src = 'https://drive-render.corp.amazon.com/view/pjbyrne@/Script%20Logos/Rocket.png'; // Replace 'your_image_url_here' with the URL of your image
        imageElement.alt = 'Rocket Image'; // Add alt text for accessibility
        imageElement.style.width = '50px'; // Adjust the width of the image as needed
        imageElement.style.marginTop = '-14px'; // Adjust the width of the image as needed


        // Append the image element
        brandElement.appendChild(imageElement);

        // Apply styling to the entire text content
        brandElement.style.fontFamily = 'Roboto, sans-serif'; // Roboto font family
        brandElement.style.fontSize = '22px'; // Larger font size
        brandElement.style.color = 'black'; // Black color
        brandElement.style.textStroke = '1px white'; // Add text stroke for 3D effect
        brandElement.style.textDecoration = 'none'; // Remove underline

        // Add hover effect
        brandElement.style.transition = 'color 0.3s, transform 0.3s'; // Smooth transition for color and transform
        brandElement.style.cursor = 'pointer'; // Change cursor to pointer on hover

        // Define hover styles
        brandElement.addEventListener('mouseover', function() {
            brandElement.style.color = 'blue'; // Change text color to blue on hover
        });

        // Define unhover styles
        brandElement.addEventListener('mouseout', function() {
            brandElement.style.color = 'black'; // Change text color back to black on unhover
        });



        ///// WIKI ELEMENT ///////
        var authorElement = document.createElement('span');
        authorElement.textContent = '';
        authorElement.style.marginLeft = '10px'; // Add more margin-left

        // Create the author link
        var authorLink = document.createElement('a');
        var authorName = '👉 WIKI'; // 
        authorLink.href = 'https://w.amazon.com/bin/view/Mobility_Maximized/'; // Embed author name at the end of URL
        authorLink.textContent = authorName;
        authorLink.target = '_blank'; // Open link in a new window/tab

        // Append the author link to the author element
        authorElement.appendChild(authorLink);

        // Apply styles to create a 3D bubble effect
        authorElement.style.padding = '5px';
        authorElement.style.borderRadius = '10px';
        authorElement.style.marginTop = '10px'; // Change the margin top to separate it from the brand element
        authorElement.style.marginRight = '5px'; // Change the margin top to separate it from the brand element
        authorElement.style.backgroundColor = '#85c1e9';
        authorElement.style.boxShadow = '2px 2px 5px rgba(0, 0, 0, 0.1)';
        authorElement.style.float = 'right'; // Align to the right side

        // Append the authorElement below the brandElement
        brandElement.insertAdjacentElement('afterend', authorElement);
    }

    // Call the function to add the maximized text and apply styling
    addMaximizedText();



    /////TESTING 💾 DOWNLOAD BUTTON ////

    // Function to update the native download link and customize the filename
    function customizeNativeDownloadLink() {
        // Locate the download link by its href containing "/part/search/download_results"
        const downloadLink = document.querySelector('a[href*="/part/search/download_results"]');

        if (!downloadLink) {
            console.error('Download link not found!');
            return;
        }

        // Replace the text content with "DOWNLOAD💾"
        downloadLink.textContent = 'DOWNLOAD💾';
        downloadLink.style.fontWeight = 'bold'; // Set bold style

        // Listen for the click event to intercept the download process
        downloadLink.addEventListener('click', (event) => {
            event.preventDefault(); // Prevent the default download action

            // Fetch the URL from the href attribute
            const downloadURL = downloadLink.href;

            // Function to format the date like 01-DEC-24
            function formatDate() {
                const monthNames = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
                const date = new Date();
                const day = String(date.getDate()).padStart(2, '0'); // Pad with leading 0 if necessary
                const month = monthNames[date.getMonth()]; // Get month name in 3-letter format
                const year = date.getFullYear().toString().slice(-2); // Get last 2 digits of the year
                return `${day}-${month}-${year}`;
            }

            // Function to get the username from the <p> element with id="hello-user"
            function getUsername() {
                const userElement = document.querySelector('#hello-user');
                if (userElement) {
                    const userText = userElement.textContent;
                    return userText.replace('Hello, ', '').trim(); // Remove "Hello, " and trim whitespace
                }
                return 'User'; // Default to 'User' if username not found
            }

            // Fetch the CSV content from the native download URL
            fetch(downloadURL)
                .then(response => response.blob())
                .then(blob => {
                // Create a downloadable link with a custom filename
                const link = document.createElement('a');
                const currentDate = formatDate(); // Use formatted date
                const username = getUsername(); // Get the username
                const fileName = `Mobility_Results_${currentDate}_${username}.csv`; // Custom file name

                // Set the blob as the href and download attribute
                link.href = URL.createObjectURL(blob);
                link.setAttribute('download', fileName);
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            })
                .catch(error => console.error('Failed to fetch the CSV file:', error));
        });
    }

    // Call the function to customize the native download link
    customizeNativeDownloadLink();












    ///////🧭 STYLE NAVBAR ///////////


    function styleNavbar() {
        const navbarDiv = document.querySelector('.navbar.navbar-fixed-top');
        if (navbarDiv) {
            navbarDiv.style.borderRadius = '10px';
            navbarDiv.style.marginRight = '5px';
            navbarDiv.style.marginLeft = '0px';
            navbarDiv.style.marginTop = '-8px';
            navbarDiv.style.width = '100%'; // Change width to 100%
            navbarDiv.style.textAlign = 'center';
            navbarDiv.style.backgroundColor = '#ccd1d1';

            navbarDiv.style.borderLeft = '4px solid #e0e0e0'; // Slightly darker color on left
            navbarDiv.style.borderRight = '4px solid #e0e0e0'; // Slightly darker color on right
            navbarDiv.style.borderBottom = '4px solid #d0d0d0'; // Slightly darker color on bottom

            // Add box shadow for lifted effect with blue light
            navbarDiv.style.boxShadow = '0px 2px 6px rgba(0, 0, 255, 0.1)'; // Reduced intensity shadow
        } else {
            console.error('Navbar element not found.');
        }
    }

    // Call the function to apply styles to the navbar
    styleNavbar();





    ///// DARK MODE SWITCH 🌑🌔////////

    // Declare switchContainer globally within the scope of your script
    var switchContainer;

    // Configuration for elements to be toggled
    var elementsConfig = [
        {
            selector: '.search_pagination', // Tool Bar
            darkBackground: '#85929e', // Dark blue background for dark mode
            darkColor: 'black', // Light grey text for dark mode
            buttonBorderColor: '#ffffff', // White border for buttons in dark mode
            dropdownBorderColor: '#ffffff', // White border for dropdowns in dark mode
            revertBackground: '#d0d3d4' // Background color to revert to when dark mode is untoggled
        },
        {
            selector: '.default-div',
            darkBackground: '#212f3d', // Different dark blue for another element
            darkColor: '#ffffff' // White text for dark mode
        },
        {
            selector: '.table thead th, .table thead th a', // TABLE HEADER CELLS
            darkBackground: '#f2f3f4', // Different dark blue for another element
            darkColor: '#212f3d', // White text for dark mode
            revertColor: '#ffffff', // Text color to revert to when dark mode is untoggled
            revertBackground: '#31465E' // Background color to revert to when dark mode is untoggled
        },
        {
            selector: '.table tbody td', // TABLE ROW CELLS
            darkBackground: '#85929e', // Different dark blue for another element
            darkColor: '#ffffff', // Dark grey text for dark mode
            revertColor: 'black' // Color to revert to when dark mode is untoggled
        },
        {
            selector: '.table td a', // SERIAL CELLS
            darkBackground: '#85929e', // Different dark blue for another element
            darkColor: '#ffffff' // Dark grey text for dark mode
        },
        {
            selector: '.search_refinements', // New element to be toggled
            darkBackground: '#85929e', // Bright blue background for dark mode
            darkColor: '#ffffff', // Optional, in case you want to change text color
            revertBackground: '#d0d3d4' // Background color to revert to when dark mode is untoggled
        }
    ];

    // Variables to store original styles for specific elements
    var originalStyles = [];

    function toggleDarkMode() {
        const isDarkMode = document.getElementById('darkModeSwitch').checked;
        localStorage.setItem('darkMode', isDarkMode); // Save the state to localStorage

        elementsConfig.forEach((config, index) => {
            const elements = document.querySelectorAll(config.selector);
            if (!originalStyles[index]) {
                originalStyles[index] = [];
            }
            elements.forEach((element, elementIndex) => {
                if (!originalStyles[index][elementIndex]) {
                    originalStyles[index][elementIndex] = {
                        backgroundColor: window.getComputedStyle(element).backgroundColor,
                        color: window.getComputedStyle(element).color,
                        buttonStyles: [],
                        dropdownStyles: []
                    };
                    element.querySelectorAll('button').forEach((button, buttonIndex) => {
                        originalStyles[index][elementIndex].buttonStyles[buttonIndex] = {
                            borderColor: window.getComputedStyle(button).borderColor
                        };
                    });
                    element.querySelectorAll('select').forEach((select, selectIndex) => {
                        originalStyles[index][elementIndex].dropdownStyles[selectIndex] = {
                            borderColor: window.getComputedStyle(select).borderColor
                        };
                    });
                }
                if (isDarkMode) {
                    // Apply dark theme styles
                    element.style.backgroundColor = config.darkBackground;
                    element.style.color = config.darkColor;
                    // Apply styles to child buttons and dropdowns
                    if (config.buttonBorderColor) {
                        element.querySelectorAll('button').forEach(button => {
                            button.style.borderColor = config.buttonBorderColor;
                        });
                    }
                    if (config.dropdownBorderColor) {
                        element.querySelectorAll('select').forEach(dropdown => {
                            dropdown.style.borderColor = config.dropdownBorderColor;
                        });
                    }
                } else {
                    // Restore the original styles
                    element.style.backgroundColor = config.revertBackground || originalStyles[index][elementIndex].backgroundColor;
                    element.style.color = config.revertColor || originalStyles[index][elementIndex].color; // Revert to black if specified, else original color
                    // Restore styles to child buttons and dropdowns
                    originalStyles[index][elementIndex].buttonStyles.forEach((buttonStyle, buttonIndex) => {
                        element.querySelectorAll('button')[buttonIndex].style.borderColor = buttonStyle.borderColor;
                    });
                    originalStyles[index][elementIndex].dropdownStyles.forEach((dropdownStyle, dropdownIndex) => {
                        element.querySelectorAll('select')[dropdownIndex].style.borderColor = dropdownStyle.borderColor;
                    });
                }
            });
        });
    }

    function setupSwitch() {
        const navbar = document.querySelector('.navbar-inner');
        if (navbar) {
            switchContainer = document.createElement('div');
            switchContainer.style.display = 'inline-block';
            switchContainer.style.verticalAlign = 'middle';
            switchContainer.style.margin = '10px';

            const switchInput = document.createElement('input');
            switchInput.type = 'checkbox';
            switchInput.id = 'darkModeSwitch';
            switchInput.style.display = 'none';

            const switchLabel = document.createElement('label');
            switchLabel.htmlFor = 'darkModeSwitch';
            switchLabel.style.position = 'relative';
            switchLabel.style.display = 'inline-block';
            switchLabel.style.width = '36px';
            switchLabel.style.height = '16px';
            switchLabel.style.backgroundColor = '#ccc';
            switchLabel.style.borderRadius = '20px';
            switchLabel.style.border = '1px solid black'; // Adding a black border around the switch
            switchLabel.style.cursor = 'pointer';

            // Set the initial title based on the current state of the switch
            switchLabel.title = switchInput.checked ? "Toggle Light Mode 🌔" : "Toggle Dark Mode 🌒";

            const sliderSpan = document.createElement('span');
            sliderSpan.style.position = 'absolute';
            sliderSpan.style.top = '1px'; // Adjust slightly for vertical centering
            sliderSpan.style.left = '2px';
            sliderSpan.style.width = '14px'; // Slightly bigger
            sliderSpan.style.height = '14px';
            sliderSpan.style.backgroundColor = 'white';
            sliderSpan.style.borderRadius = '50%';
            sliderSpan.style.transition = '.4s';
            sliderSpan.style.display = 'flex';
            sliderSpan.style.alignItems = 'center';
            sliderSpan.style.justifyContent = 'center';
            sliderSpan.style.fontSize = '12px'; // Make emoji bigger
            sliderSpan.style.userSelect = 'none';
            sliderSpan.textContent = '☀️'; // Default emoji


            switchLabel.appendChild(sliderSpan);
            switchContainer.appendChild(switchInput);
            switchContainer.appendChild(switchLabel);
            navbar.appendChild(switchContainer);

            switchInput.addEventListener('change', toggleDarkMode);
            switchInput.addEventListener('change', function() {
                if (this.checked) {
                    switchLabel.style.backgroundColor = '#2196F3'; // Blue when active
                    sliderSpan.style.transform = 'translateX(20px)';
                    sliderSpan.textContent = '🌙'; // Show moon
                    switchLabel.title = "Toggle Light Mode 🌔";
                } else {
                    switchLabel.style.backgroundColor = '#ccc'; // Grey when inactive
                    sliderSpan.style.transform = 'translateX(0)';
                    sliderSpan.textContent = '☀️'; // Show sun
                    switchLabel.title = "Toggle Dark Mode 🌒";
                }
            });


            // Load the saved state from localStorage
            const savedDarkMode = localStorage.getItem('darkMode') === 'true';
            switchInput.checked = savedDarkMode;
            if (savedDarkMode) {
                switchLabel.style.backgroundColor = '#2196F3';
                sliderSpan.style.transform = 'translateX(20px)';
                sliderSpan.textContent = '🌙'; // Moon emoji when loading
                switchLabel.title = "Toggle Light Mode 🌔";
                toggleDarkMode();
            }

        }
    }

    // Call setup function
    setupSwitch();







    /////🧰 TOOL CONTAINER ////
    const formContainer = document.createElement('div');
    formContainer.id = 'formContainer';

    // Append the form container to the document body
    document.body.appendChild(formContainer);







    //////🚩 ADD FLAGS BASED ON CLUSTER CODE 🚩//////
    // Function to add flag based on cluster or site code
    function addFlagToCluster() {
        const headings = document.querySelectorAll('.search_refinements h5');
        headings.forEach(heading => {
            const headingText = heading.textContent.trim();
            if (headingText === 'Cluster Name' || headingText === 'Site') {
                const items = heading.nextElementSibling.querySelectorAll('li');
                items.forEach(item => {
                    // Select both hyperlinked and non-hyperlinked elements
                    const hyperlinked = item.querySelector('a');
                    const code = hyperlinked ? hyperlinked.textContent.trim() : item.textContent.trim();
                    let flag = '';
                    // Add flags based on codes
                    if (code.includes('ARN')) {
                        flag = '🇸🇪'; // Sweden flag (Stockholm Arlanda Airport)
                    } else if (code.includes('BAH')) {
                        flag = '🇧🇭'; // Bahrain flag (Bahrain International Airport)
                    } else if (code.includes('CDG')) {
                        flag = '🇫🇷'; // France flag (Charles de Gaulle Airport)
                    } else if (code.includes('CMH')) {
                        flag = '🇺🇸'; // France flag (Charles de Gaulle Airport)
                    } else if (code.includes('CPT')) {
                        flag = '🇿🇦'; // South Africa flag (Cape Town International Airport)
                    } else if (code.includes('DUB')) {
                        flag = '🇮🇪'; // Ireland flag (Dublin Airport)
                    } else if (code.includes('LHR')) {
                        flag = '🇬🇧'; // UK flag (London Heathrow Airport)
                    } else if (code.includes('MXP')) {
                        flag = '🇮🇹'; // Italy flag (Milan Malpensa Airport)
                    } else if (code.includes('ZAZ')) {
                        flag = '🇪🇸'; // Spain flag (Zaragoza Airport)
                    }
                    else if (code.includes('MAD')) {
                        flag = '🇪🇸'; // Spain flag (Zaragoza Airport)
                    }

                    else if (code.includes('AMS')) {
                        flag = '🇳🇱'; // Netherlands flag
                    }
                    else if (code.includes('ZRH')) {
                        flag = '🇨🇭'; // Switzerland flag (Zurich Airport)
                    } else if (code.includes('TLV')) {
                        flag = '🇮🇱'; // Israel flag (Ben Gurion Airport)
                    } else if (code.includes('DXB')) {
                        flag = '🇦🇪'; // UAE flag (Dubai International Airport)
                    } else if (code.includes('MCT')) {
                        flag = '🇴🇲'; // Oman flag (Muscat International Airport)
                    } else if (code.includes('OSL')) {
                        flag = '🇳🇴'; // Norway flag (Oslo Gardermoen Airport)
                    } else if (code.includes('JNB')) {
                        flag = '🇿🇦'; // South Africa flag (O. R. Tambo International Airport)
                    } else if (code.includes('NBO')) {
                        flag = '🇰🇪'; // Kenya flag (Nairobi Jomo Kenyatta International Airport)
                    } else if (code.includes('IAD')) {
                        flag = '🇺🇸'; // USA flag (Washington Dulles International Airport)
                    } else if (code.includes('PDX')) {
                        flag = '🇺🇸'; // USA flag (Portland International Airport)
                    } else if (code.includes('NRT')) {
                        flag = '🇯🇵'; // Japan flag (Narita International Airport)
                    } else if (code.includes('SIN')) {
                        flag = '🇸🇬'; // Singapore flag (Singapore Changi Airport)
                    } else if (code.includes('SFO')) {
                        flag = '🇺🇸'; // USA flag (San Francisco International Airport)
                    } else if (code.includes('FRA')) {
                        flag = '🇩🇪'; // Germany flag (Frankfurt Airport)
                    } else if (code.includes('CVG')) {
                        flag = '🇺🇸'; // USA flag (Cincinnati/Northern Kentucky International Airport)
                    } else if (code.includes('GRU')) {
                        flag = '🇧🇷'; // Brazil flag (São Paulo/Guarulhos–Governador André Franco Montoro International Airport)
                    } else if (code.includes('SYD')) {
                        flag = '🇦🇺'; // Australia flag (Sydney Airport)
                    } else if (code.includes('POP')) {
                        flag = ''; // Dominican Republic flag (Puerto Plata Airport)
                    } else if (code.includes('BOM')) {
                        flag = '🇮🇳'; // India flag (Chhatrapati Shivaji Maharaj International Airport)
                    } else if (code.includes('DCA')) {
                        flag = '🇺🇸'; // USA flag (Ronald Reagan Washington National Airport)
                    } else if (code.includes('ICN')) {
                        flag = '🇰🇷'; // South Korea flag (Incheon International Airport)
                    } else if (code.includes('YUL')) {
                        flag = '🇨🇦'; // Canada flag (Montréal-Pierre Elliott Trudeau International Airport)
                    } else if (code.includes('SEA')) {
                        flag = '🇺🇸'; // USA flag (Seattle-Tacoma International Airport)
                    } else if (code.includes('BJS')) {
                        flag = '🇨🇳'; // China flag (Beijing Capital International Airport)
                    } else if (code.includes('ZHY')) {
                        flag = '🇨🇳'; // China flag (Zhengzhou Xinzheng International Airport)
                    } else if (code.includes('LCK')) {
                        flag = '🇺🇸'; // USA flag (Rickenbacker International Airport)
                    } else if (code.includes('KIX')) {
                        flag = '🇯🇵'; // Japan flag (Kansai International Airport)
                    } else if (code.includes('HYD')) {
                        flag = '🇮🇳'; // India flag (Rajiv Gandhi International Airport)
                    } else if (code.includes('DEN')) {
                        flag = '🇺🇸'; // USA flag (Denver International Airport)
                    } else if (code.includes('CGK')) {
                        flag = '🇮🇩'; // Indonesia flag (Soekarno–Hatta International Airport)
                    } else if (code.includes('PEK')) {
                        flag = '🇨🇳'; // China flag (Beijing Capital International Airport)
                    } else if (code.includes('MEL')) {
                        flag = '🇦🇺'; // Australia flag (Melbourne Airport)
                    } else if (code.includes('YYC')) {
                        flag = '🇨🇦'; // Canada flag (Calgary International Airport)
                    } else if (code.includes('JFK')) {
                        flag = '🇺🇸'; 
                    } else if (code.includes('ATL')) {
                        flag = '🇺🇸'; // USA flag (Hartsfield–Jackson Atlanta International Airport)
                    } else if (code.includes('SAT')) {
                        flag = '🇺🇸'; // USA flag (San Antonio International Airport)
                    } else if (code.includes('FFZ')) {
                        flag = '🇺🇸'; // Unknown airport
                    } else if (code.includes('KUL')) {
                        flag = '🇲🇾'; // Malaysia flag (Kuala Lumpur International Airport)
                    } else if (code.includes('LAX')) {
                        flag = '🇺🇸'; // USA flag (Los Angeles International Airport)
                    } else if (code.includes('BWI')) {
                        flag = '🇺🇸'; // USA flag (Baltimore/Washington International Thurgood Marshall Airport)
                    } else if (code.includes('PHX')) {
                        flag = '🇺🇸'; // USA flag (Phoenix Sky Harbor International Airport)
                    } else if (code.includes('LUX')) {
                        flag = '🇱🇺'; // Luxembourg flag (Luxembourg Airport)
                    } else if (code.includes('SJC')) {
                        flag = '🇺🇸'; // USA flag (Norman Y. Mineta San José International Airport)
                    } else if (code.includes('ALE')) {
                        flag = '🇺🇸'; // Unknown airport
                    } else if (code.includes('MDW')) {
                        flag = '🇺🇸'; // USA flag (Chicago Midway International Airport)
                    } else if (code.includes('BKK')) {
                        flag = '🇹🇭'; // Thailand flag (Suvarnabhumi Airport)
                    } else if (code.includes('BPM')) {
                        flag = '🇮🇳'; // Unknown airport
                    } else if (code.includes('HKG')) {
                        flag = '🇭🇰'; // Hong Kong flag (Hong Kong International Airport)
                    } else if (code.includes('PDT')) {
                        flag = '🇺🇸'; // Unknown airport
                    } else if (code.includes('LTW')) {
                        flag = '🇺🇸'; // Unknown airport
                    } else if (code.includes('FKE')) {
                        flag = ''; // Unknown airport
                    } else if (code.includes('RAD')) {
                        flag = '🇻🇬'; // Unknown airport
                    } else if (code.includes('BER')) {
                        flag = '🇩🇪'; // Unknown airport
                    } else if (code.includes('OSU')) {
                        flag = '🇺🇸'; // Unknown airport
                    } else if (code.includes('APA')) {
                        flag = '🇺🇸'; // Unknown airport
                    } else if (code.includes('QRO')) {
                        flag = '🇲🇽'; // Unknown airport
                    } else if (code.includes('SCL')) {
                        flag = '🇨🇱'; // Chile flag (Arturo Merino Benítez International Airport)
                    } else if (code.includes('BOG')) {
                        flag = '🇨🇴'; // Colombia flag (El Dorado International Airport)
                    } else if (code.includes('EZE')) {
                        flag = '🇦🇷'; // Argentina flag (Ministro Pistarini International Airport)
                    } else if (code.includes('OMD')) {
                        flag = ''; // Unknown airport
                    } else if (code.includes('AKL')) {
                        flag = '🇳🇿'; // New Zealand flag (Auckland Airport)
                    }

                    // Prepend flag to code
                    if (flag !== '') {
                        const flagSpan = document.createElement('span');
                        flagSpan.textContent = flag + ' ';
                        flagSpan.style.fontSize = '17px'; // Increase font size
                        item.insertBefore(flagSpan, item.firstChild);
                    }
                });
            }
        });
    }

    // Call the function to add flags based on cluster or site code
    addFlagToCluster();






    /////☰ SIDE MENU HEADING EMOJIS /////////
    const headings = document.querySelectorAll('.search_refinements h5');

    headings.forEach((heading) => {
        const headingName = heading.textContent.trim();
        let emoji = '';
        switch (headingName) {
            case 'Cluster Name':
                emoji = '🖧';
                break;
            case 'Site':
                emoji = '🏬';
                break;
            case 'Room':
                emoji = '🚪';
                break;
            case 'Bin':
                emoji = '🗑️';
                break;
            case 'Model':
                emoji = '🆔';
                break;
            case 'Model Apn':
                emoji = '🆔';
                break;
            case 'Model Mpn':
                emoji = '🆔';
                break;
            case 'Unit Cost':
                emoji = '💲';
                break;
            case 'Vendor':
                emoji = '🛒';
                break;
            case 'Type Name':
                emoji = '⚙️';
                break;
            case 'State':
                emoji = '📑';
                break;
            case 'Abstract State':
                emoji = '📝';
                break;
            case 'Category':
                emoji = '🗂️';
                break;
            case 'Tag':
                emoji = '🏷️';
                break;
            default:
                emoji = '';
                break;
        }
        if (emoji !== '') {
            const emojiSpan = document.createElement('span');
            emojiSpan.textContent = emoji + ' '; // Add space after emoji
            emojiSpan.style.fontSize = '14px'; // Increase font size
            heading.insertBefore(emojiSpan, heading.firstChild);
        }
    });



    /// 📦 TOTE ID HEADING //////

    function toteIdHeader() {
        console.log("📦 TOTE ID toteIdHeader function is running");

        var targetDiv = document.querySelector('.search_refinements.span');

        if (targetDiv) {
            var newHeading = document.createElement('h5');
            newHeading.innerHTML = '<span style="font-size: 14px;">📦 </span>Tote ID';

            var newList = document.createElement('ul');
            newList.className = 'linklist';

            targetDiv.appendChild(newHeading);
            targetDiv.appendChild(newList);

            console.log("📦 TOTE ID Elements successfully appended to the target div");

            fetchAndDisplayToteIds(newList);
        } else {
            console.error("📦 TOTE ID Target div not found");
        }
    }

    async function fetchAndDisplayToteIds(newList) {
        // Fetch the literal URL of the current page from the URL bar
        let currentUrl = window.location.href;
        console.log('📦 TOTE ID Using current page URL:', currentUrl);

        // Replace the max_rows parameter with 10000
        currentUrl = currentUrl.replace(/max_rows=\d+/i, 'max_rows=10000');
        console.log('📦 TOTE ID Modified URL:', currentUrl);

        try {
            const response = await fetch(currentUrl, {
                method: 'GET',
                headers: {
                    'Content-Type': 'text/html',
                },
            });

            if (!response.ok) {
                console.error('📦 TOTE ID Failed to fetch the page:', response.statusText);
                return;
            }

            const htmlText = await response.text();
            console.log('📦 TOTE ID Fetched HTML response length:', htmlText.length);

            const parser = new DOMParser();
            const doc = parser.parseFromString(htmlText, 'text/html');

            const serialIds = doc.querySelectorAll('h3'); // Get all <h3> tags
            const toteIds = {};
            const toteToSerialsMap = {};

            console.log("📦 TOTE ID Parsing Serial and Tote ID pairs:");

            serialIds.forEach((serialIdElement, index) => {
                const serialId = serialIdElement.textContent.replace('Serial ID: ', '').trim();
                let nextElement = serialIdElement.nextElementSibling;
                let foundToteId = false;

                // Traverse the siblings of the current serial ID to find the Tote ID
                while (nextElement && nextElement.tagName !== 'H3') {
                    if (nextElement.textContent.includes('Tote Asset ID:')) {
                        const toteId = nextElement.textContent.replace('Tote Asset ID:', '').trim();
                        if (toteId) {
                            if (!toteToSerialsMap[toteId]) {
                                toteToSerialsMap[toteId] = [];
                            }
                            toteToSerialsMap[toteId].push(serialId);
                            toteIds[toteId] = (toteIds[toteId] || 0) + 1;

                            console.log(`📦 Serial ID: ${serialId} -> Tote ID: ${toteId}`);
                            foundToteId = true;
                        }
                        break; // Stop once we find the Tote ID
                    }
                    nextElement = nextElement.nextElementSibling;
                }

                if (!foundToteId) {
                    console.log(`❌ Serial ID: ${serialId} -> No Tote ID found`);
                }
            });

            console.log('📦 TOTE ID Total unique Tote IDs found:', Object.keys(toteIds).length);

            newList.innerHTML = '';

            const shouldHyperlink = Object.keys(toteIds).length > 1;

            // Display each Tote ID with its count and a hyperlink if needed
            for (const [toteId, count] of Object.entries(toteIds)) {
                const listItem = document.createElement('li');

                if (shouldHyperlink) {
                    // Join serials with a + for the URL, but ensure they appear as spaces in the search box
                    const serialsForToteId = toteToSerialsMap[toteId].join(' ');
                    const linkUrl = `https://mobility.AMAZON.com/part/search?search_type=all&search_string=${encodeURIComponent(serialsForToteId).replace(/%20/g, '+')}&max_rows=10000&query=GO`;

                    const link = document.createElement('a');
                    link.href = linkUrl;
                    link.textContent = toteId;
                    link.target = "_blank";

                    listItem.appendChild(link);
                } else {
                    listItem.textContent = toteId;
                }

                const countSpan = document.createElement('span');
                countSpan.textContent = ` (${count})`;
                countSpan.style.color = 'black';

                listItem.appendChild(countSpan);
                newList.appendChild(listItem);
            }

            console.log("📦 TOTE ID Tote IDs successfully appended to the list.");

        } catch (error) {
            console.error('📦 TOTE ID Error fetching or processing the data:', error);
        }
    }

    // Ensure the function is called after it is defined
    toteIdHeader();




    //// TESTING - PO NUMBER FILTER //////


    function poNumberHeader() {
        console.log("🧾 PO NUMBER poNumberHeader function is running");

        var targetDiv = document.querySelector('.search_refinements.span');

        if (targetDiv) {
            var newHeading = document.createElement('h5');
            newHeading.innerHTML = '<span style="font-size: 14px;">🧾 </span>Po Number';

            var newList = document.createElement('ul');
            newList.className = 'linklist';

            targetDiv.appendChild(newHeading);
            targetDiv.appendChild(newList);

            console.log("🧾 PO NUMBER Elements successfully appended to the target div");

            fetchAndDisplayPoNumbers(newList);
        } else {
            console.error("🧾 PO NUMBER Target div not found");
        }
    }

    async function fetchAndDisplayPoNumbers(newList) {
        // Define flag mappings for different substrings
        const flagMappings = {
            'BH': '🇧🇭', // Flag emoji for Bahrain
            'US': '🇺🇸', // Flag emoji for USA
            'ZA': '🇿🇦', // Flag emoji for South Africa
            'AE': '🇦🇪', // Flag emoji for UAE
            'IT': '🇮🇹', // Flag emoji for Italy
            'DE': '🇩🇪', // Flag emoji for Germany
            'SW': '🇸🇪', // Flag emoji for Sweden
            'LHR': '🇬🇧', // Flag emoji for UK
            'ZHY': '🇨🇳', // Flag emoji for China
            'EU': '🇮🇪', // Flag emoji for Ireland
            'FR': '🇫🇷', // Flag emoji for France
            'ES': '🇪🇸', // Flag emoji for Spain
            'CH': '🇨🇭', // Flag emoji for Switzerland
            'IL': '🇮🇱', // Flag emoji for Israel
            'AU': '🇦🇺', // Flag emoji for Australia
            'CA': '🇨🇦', // Flag emoji for Canada
            'BR': '🇧🇷' // Flag emoji for Brazil
            // Add more mappings for other substrings as needed
        };

        // Fetch the literal URL of the current page from the URL bar
        let currentUrl = window.location.href;
        console.log('🧾 PO NUMBER Using current page URL:', currentUrl);

        // Replace the max_rows parameter with 10000
        currentUrl = currentUrl.replace(/max_rows=\d+/i, 'max_rows=10000');
        console.log('🧾 PO NUMBER Modified URL:', currentUrl);

        try {
            const response = await fetch(currentUrl, {
                method: 'GET',
                headers: {
                    'Content-Type': 'text/html',
                },
            });

            if (!response.ok) {
                console.error('🧾 PO NUMBER Failed to fetch the page:', response.statusText);
                return;
            }

            const htmlText = await response.text();
            console.log('🧾 PO NUMBER Fetched HTML response length:', htmlText.length);

            const parser = new DOMParser();
            const doc = parser.parseFromString(htmlText, 'text/html');

            const serialIds = doc.querySelectorAll('h3'); // Get all <h3> tags
            const poNumbers = {};
            const poToSerialsMap = {};

            console.log("🧾 PO NUMBER Parsing Serial and Po Number pairs:");

            serialIds.forEach((serialIdElement, index) => {
                const serialId = serialIdElement.textContent.replace('Serial ID: ', '').trim();
                let nextElement = serialIdElement.nextElementSibling;
                let foundPoNumber = false;

                // Traverse the siblings of the current serial ID to find the Po Number
                while (nextElement && nextElement.tagName !== 'H3') {
                    if (nextElement.textContent.includes('PO Number:')) {
                        const poNumber = nextElement.textContent.replace('PO Number:', '').trim();
                        if (poNumber) {
                            if (!poToSerialsMap[poNumber]) {
                                poToSerialsMap[poNumber] = [];
                            }
                            poToSerialsMap[poNumber].push(serialId);
                            poNumbers[poNumber] = (poNumbers[poNumber] || 0) + 1;

                            console.log(`🧾 Serial ID: ${serialId} -> Po Number: ${poNumber}`);
                            foundPoNumber = true;
                        }
                        break; // Stop once we find the Po Number
                    }
                    nextElement = nextElement.nextElementSibling;
                }

                if (!foundPoNumber) {
                    console.log(`❌ Serial ID: ${serialId} -> No Po Number found`);
                }
            });

            console.log('🧾 PO NUMBER Total unique Po Numbers found:', Object.keys(poNumbers).length);

            newList.innerHTML = '';

            const shouldHyperlink = Object.keys(poNumbers).length > 1;

            // Display each Po Number with its count and a hyperlink if needed
            for (const [poNumber, count] of Object.entries(poNumbers)) {
                const listItem = document.createElement('li');

                // Determine the flag to display based on the first 2-3 characters of the Po Number
                const poPrefix3 = poNumber.substring(0, 3).toUpperCase();
                const poPrefix2 = poNumber.substring(0, 2).toUpperCase();
                const flag = flagMappings[poPrefix3] || flagMappings[poPrefix2] || '';

                if (shouldHyperlink) {
                    // Join serials with a + for the URL, but ensure they appear as spaces in the search box
                    const serialsForPoNumber = poToSerialsMap[poNumber].join(' ');
                    const linkUrl = `https://mobility.AMAZON.com/part/search?search_type=all&search_string=${encodeURIComponent(serialsForPoNumber).replace(/%20/g, '+')}&max_rows=10000&query=GO`;

                    const link = document.createElement('a');
                    link.href = linkUrl;
                    link.textContent = `${flag} ${poNumber}`;
                    link.target = "_blank";

                    listItem.appendChild(link);
                } else {
                    listItem.textContent = `${flag} ${poNumber}`;
                }

                const countSpan = document.createElement('span');
                countSpan.textContent = ` (${count})`;
                countSpan.style.color = 'black';

                listItem.appendChild(countSpan);
                newList.appendChild(listItem);
            }

            console.log("🧾 PO NUMBER Po Numbers successfully appended to the list.");

        } catch (error) {
            console.error('🧾 PO NUMBER Error fetching or processing the data:', error);
        }
    }

    // Ensure the function is called after it is defined
    poNumberHeader();








    ////// 🔝 STYLE MOBILITY TABLE HEADER ////////
    function styleTable() {
        // Find the table element using XPath
        const table = document.evaluate("/html/body/div[3]/div[1]/div[2]/section/div[3]/table", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

        // Check if the table element exists
        if (table) {
            // Apply styles to the table
            table.style.backgroundColor = "#f8f9f9"; // Background color #FF9900 #35465c
            table.style.fontSize = "12px";
            table.style.borderRadius = "10px"; // Curved edges
        }
    }

    // Call the function to style the table
    styleTable();








    ////🙂 ADD EMOJIS TO COLUMN HEADERS ///////


    function addEmojisToColumnHeaders() {
        const columnHeaders = document.querySelectorAll('.table th a');

        columnHeaders.forEach((header) => {
            const headerText = header.textContent.trim();
            let emoji = '';
            switch (headerText) {
                case 'Part ID':
                    emoji = '🆔';
                    break;
                case 'Serial ID':
                    emoji = '▌│█║▌║▌║';
                    break;
                case 'Tags':
                    emoji = '🏷️';
                    break;
                case 'Shipment Id':
                    emoji = '🚚';
                    break;
                case 'Shipment Type':
                    emoji = '📦';
                    break;
                case 'Cluster':
                    emoji = '🌐';
                    break;
                case 'Site':
                    emoji = '🏢';
                    break;
                case 'Room':
                    emoji = '🚪';
                    break;
                case 'Bin':
                    emoji = '🗑️';
                    break;
                case 'Model':
                    emoji = '🛠️';
                    break;
                case 'Model MPN':
                    emoji = '📝';
                    break;
                case 'Model APN':
                    emoji = '📝';
                    break;
                case 'Model Description':
                    emoji = '';
                    break;
                case 'Unit Cost':
                    emoji = '💲';
                    break;
                case 'Vendor':
                    emoji = '🛒';
                    break;
                case 'Type':
                    emoji = '⚙️';
                    break;
                case 'State':
                    emoji = '📑';
                    break;
                case 'Abstract State':
                    emoji = '📝';
                    break;
                case 'Category':
                    emoji = '🗂️';
                    break;
                case 'Part Reserved Status':
                    emoji = '🚫';
                    break;
                case 'User Custody':
                    emoji = '👤';
                    break;
                case 'External Serial ID':
                    emoji = '🔗';
                    break;
                case 'Tote Asset ID':
                    emoji = '🛍️';
                    break;
                case 'PO Number':
                    emoji = '📋';
                    break;
                case 'RMA ID':
                    emoji = '📦';
                    break;
                case 'Outbound RMA ID':
                    emoji = '📤';
                    break;
                case 'Tracking ID':
                    emoji = '🚛';
                    break;
                case 'Outbound Tracking ID':
                    emoji = '🚚';
                    break;
                case 'Pallet Asset ID':
                    emoji = '📦';
                    break;
                case 'Build ID':
                    emoji = '🛠️';
                    break;
                case 'Repair ID':
                    emoji = '🔧';
                    break;
                case 'Last Updated Time':
                    emoji = '⏰';
                    break;
                case 'First Recovered Time':
                    emoji = '🕒';
                    break;
                case 'Received At':
                    emoji = '📅';
                    break;
                case 'Last Known Location':
                    emoji = '📍';
                    break;
                case 'Secure Bin Serial ID':
                    emoji = '🔒';
                    break;
                case 'Smart Device Serial ID':
                    emoji = '📱';
                    break;
                case 'Smart Device Type':
                    emoji = '📱';
                    break;
                case 'Smart Device Location':
                    emoji = '📍';
                    break;
                default:
                    emoji = '';
                    break;
            }

            if (emoji !== '') {
                const emojiSpan = document.createElement('span');
                emojiSpan.textContent = ' ' + emoji; // Add space before emoji
                emojiSpan.style.fontSize = '12px'; // Adjust font size
                header.appendChild(emojiSpan); // Append emoji span to header
            }
        });
    }

    // Call the function to add emojis to column headers
    // addEmojisToColumnHeaders(); // Commented out to disable the function temporarily






    ///// 🫣👀 HIDE COLUMNS

    // Function to add a button to the webpage for toggling column visibility
    const addColumnToggleButton = () => {
        // Function to apply column visibility preferences
        const applyPreferences = () => {
            const columnIndices = Array.from({ length: 40 }, (_, i) => i + 1);
            columnIndices.forEach(index => {
                const savedPreference = localStorage.getItem(`column_visibility_${index}`);
                if (savedPreference !== null) {
                    const columns = document.querySelectorAll(`table.table-bordered.table-striped.table-condensed tbody tr td:nth-child(${index}), table.table-bordered.table-striped.table-condensed thead th:nth-child(${index})`);
                    columns.forEach(column => {
                        column.style.display = savedPreference === 'true' ? '' : 'none';
                    });
                }
            });
        };

        // Apply preferences on page load
        applyPreferences();

        // Function to check the status of Toggle All checkbox
        const checkToggleAllStatus = () => {
            const allChecked = columnIndices.every(index => {
                const savedPreference = localStorage.getItem(`column_visibility_${index}`);
                return savedPreference === 'true';
            });
            toggleAllInput.checked = allChecked;
        };

        // Create the slider window
        const sliderWindow = document.createElement('div');
        sliderWindow.style.position = 'fixed';
        sliderWindow.style.top = '60%';
        sliderWindow.style.left = '40%';
        sliderWindow.style.transform = 'translate(-50%, -50%)';
        sliderWindow.style.padding = '20px';
        sliderWindow.style.backgroundColor = '#f4d03f'; // Background color as specified
        sliderWindow.style.border = '2px solid black'; // Border color set to black
        sliderWindow.style.borderRadius = '10px';
        sliderWindow.style.boxShadow = '0 0 10px rgba(0, 0, 0, 0.1)';
        sliderWindow.style.zIndex = '1000';
        sliderWindow.style.maxHeight = '60%';
        sliderWindow.style.overflowY = 'auto';
        sliderWindow.style.display = 'none'; // Initially hidden
        document.body.appendChild(sliderWindow);

        // Create title for the slider window
        const title = document.createElement('div');
        title.innerText = '⚙️ Column Display Preferences';
        title.style.backgroundColor = ' #2e4053 ';
        title.style.color = 'white';
        title.style.padding = '0px';
        title.style.textAlign = 'center';
        title.style.borderTopLeftRadius = '10px';
        title.style.borderTopRightRadius = '10px';
        title.style.marginBottom = '5px';
        sliderWindow.appendChild(title);

        // Column indices to include in the sliders
        const columnIndices = Array.from({ length: 40 }, (_, i) => i + 1);
        const columnHeaders = [
            '1. ☑ SELECT', '2. 📦 PART ID', '3. 🔢 SERIAL ID', '4. 🏷️ TAGS', '5. 📦 SHIPMENT ID', '6. 📦 SHIPMENT TYPE', '7. 🏭 CLUSTER', '8. 📍 SITE', '9. 🏠 ROOM', '10. 🗄️ BIN',
            '11. 📄 MODEL', '12. 🔢 MODEL MPN', '13. 🔢 MODEL APN', '14. 📄 MODEL DESCRIPTION', '15. 💲 UNIT COST', '16. 🏭 VENDOR', '17. 📄 TYPE', '18. 📄 STATE',
            '19. 📄 ABSTRACT STATE', '20. 🗂️ CATEGORY', '21. 📄 PART RESERVED STATUS', '22. 👤 USER CUSTODY', '23. 🔢 EXTERNAL SERIAL ID', '24. 📦 TOTE ASSET ID',
            '25. 📄 PO NUMBER', '26. 📄 RMA ID', '27. 📄 OUTBOUND RMA ID', '28. 🔗 TRACKING ID', '29. 🔗 OUTBOUND TRACKING ID', '30. 🗂️ PALLET ASSET ID',
            '31. 🏗️ BUILD ID', '32. 🔧 REPAIR ID', '33. 🕒 LAST UPDATED TIME', '34. 🕒 FIRST RECOVERED TIME', '35. 📥 RECEIVED AT',
            '36. 📍 LAST KNOWN LOCATION', '37. 🔒 SECURE BIN SERIAL ID', '38. 🔢 SMART DEVICE SERIAL ID', '39. 🔧 SMART DEVICE TYPE',
            '40. 📍 SMART DEVICE LOCATION'
        ];

        // Create "Toggle All" switch
        const toggleAllContainer = document.createElement('div');
        toggleAllContainer.style.display = 'flex';
        toggleAllContainer.style.alignItems = 'center';
        toggleAllContainer.style.marginBottom = '10px';
        toggleAllContainer.style.borderBottom = '1px solid black'; // Add border

        const toggleAllLabel = document.createElement('label');
        toggleAllLabel.innerText = 'ALL';
        toggleAllLabel.style.marginRight = '10px';
        toggleAllLabel.style.width = '150px'; // Adjust the width as needed
        toggleAllContainer.appendChild(toggleAllLabel);

        const toggleAllInput = document.createElement('input');
        toggleAllInput.type = 'checkbox';
        toggleAllInput.classList.add('switch');
        toggleAllContainer.appendChild(toggleAllInput);

        sliderWindow.appendChild(toggleAllContainer);

        // Event listener for "Toggle All" switch
        toggleAllInput.addEventListener('change', (event) => {
            const checked = event.target.checked;
            columnIndices.forEach(index => {
                const columns = document.querySelectorAll(`table.table-bordered.table-striped.table-condensed tbody tr td:nth-child(${index}), table.table-bordered.table-striped.table-condensed thead th:nth-child(${index})`);
                columns.forEach(column => {
                    column.style.display = checked ? '' : 'none';
                });
                // Save the preference in localStorage
                localStorage.setItem(`column_visibility_${index}`, checked);
                // Update individual switches
                const individualSwitch = sliderWindow.querySelector(`input[data-column-index="${index}"]`);
                if (individualSwitch) {
                    individualSwitch.checked = checked;
                }
            });
        });

        // Create switches for each specified column
        columnIndices.forEach((index, i) => {
            const sliderContainer = document.createElement('div');
            sliderContainer.style.display = 'flex';
            sliderContainer.style.alignItems = 'center';
            sliderContainer.style.marginBottom = '10px';
            sliderContainer.style.borderBottom = '1px solid black'; // Add border

            const label = document.createElement('label');
            label.innerText = columnHeaders[i]; // Placeholder text for column headers
            label.style.marginRight = '10px';
            label.style.width = '150px'; // Adjust the width as needed
            sliderContainer.appendChild(label);

            const input = document.createElement('input');
            input.type = 'checkbox';
            input.classList.add('switch-input');
            input.checked = true;
            input.dataset.columnIndex = index;

            const switchLabel = document.createElement('label');
            switchLabel.classList.add('switch');
            switchLabel.appendChild(input);

            const span = document.createElement('span');
            span.classList.add('switch-slider');
            switchLabel.appendChild(span);

            sliderContainer.appendChild(switchLabel);

            sliderWindow.appendChild(sliderContainer);

            // Event listener for switch change to toggle column visibility
            input.addEventListener('change', (event) => {
                const columnIndex = event.target.dataset.columnIndex;
                const checked = event.target.checked;
                const columns = document.querySelectorAll(`table.table-bordered.table-striped.table-condensed tbody tr td:nth-child(${columnIndex}), table.table-bordered.table-striped.table-condensed thead th:nth-child(${columnIndex})`);
                columns.forEach(column => {
                    column.style.display = checked ? '' : 'none';
                });
                // Save the preference in localStorage
                localStorage.setItem(`column_visibility_${columnIndex}`, checked);
                // Update Toggle All status
                checkToggleAllStatus();
            });

            // Load the preference from localStorage
            const savedPreference = localStorage.getItem(`column_visibility_${index}`);
            if (savedPreference !== null) {
                input.checked = savedPreference === 'true';
                const columns = document.querySelectorAll(`table.table-bordered.table-striped.table-condensed tbody tr td:nth-child(${index}), table.table-bordered.table-striped.table-condensed thead th:nth-child(${index})`);
                columns.forEach(column => {
                    column.style.display = savedPreference === 'true' ? '' : 'none';
                });
            }
        });

        // Check the status of Toggle All on window open
        checkToggleAllStatus();

        // Close button for the slider window
        const closeButton = document.createElement('button');
        closeButton.innerText = 'Close';
        closeButton.style.marginTop = '10px';
        closeButton.style.padding = '10px 20px';
        closeButton.style.backgroundColor = '#e74c3c';
        closeButton.style.color = '#fff';
        closeButton.style.border = 'none';
        closeButton.style.borderRadius = '5px';
        closeButton.style.cursor = 'pointer';
        closeButton.style.position = 'sticky';
        closeButton.style.bottom = '0';
        closeButton.style.width = '100%';
        closeButton.style.textAlign = 'center';
        sliderWindow.appendChild(closeButton);

        // Event listener for close button
        closeButton.addEventListener('click', () => {
            sliderWindow.style.display = 'none';
        });

        // Return a function to toggle the visibility of the slider window
        return () => {
            sliderWindow.style.display = sliderWindow.style.display === 'none' ? 'block' : 'none';
        };
    };

    // Create the column toggle button and get the toggle function
    const toggleColumnWindow = addColumnToggleButton();

    // Create the button element
    const ColumnButton = document.createElement('button');
    ColumnButton.innerText = '⚙️';
    ColumnButton.title = 'Column Display Preferences';
    ColumnButton.style.fontSize = '24px'; // Increase font size for the emoji
    ColumnButton.style.padding = '0'; // No padding
    ColumnButton.style.border = 'none'; // Remove border
    ColumnButton.style.background = 'none'; // Remove background
    ColumnButton.style.cursor = 'pointer'; // Add pointer cursor

    // Add hover effect to grow the emoji
    ColumnButton.addEventListener('mouseenter', () => {
        ColumnButton.style.transform = 'scale(1.2)';
    });
    ColumnButton.addEventListener('mouseleave', () => {
        ColumnButton.style.transform = 'scale(1)';
    });

    ColumnButton.addEventListener('click', toggleColumnWindow);
    document.body.appendChild(ColumnButton);

    // CSS for switch style
    const switchStyle = document.createElement('style');
    switchStyle.innerHTML = `
.switch {
    position: relative;
    display: inline-block;
    width: 34px;
    height: 20px;
}

.switch-input {
    opacity: 0;
    width: 0;
    height: 0;
}

.switch-slider {
    position: absolute;
    cursor: pointer;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background-color: #ccc;
    transition: .4s;
    border-radius: 20px;
}

.switch-slider:before {
    position: absolute;
    content: "";
    height: 12px;
    width: 12px;
    left: 4px;
    bottom: 4px;
    background-color: white;
    transition: .4s;
    border-radius: 50%;
}

.switch-input:checked + .switch-slider {
    background-color: #2196F3;
}

.switch-input:checked + .switch-slider:before {
    transform: translateX(14px);
}
`;
    document.head.appendChild(switchStyle);







    ////// 📉 DESCRIPTION SHORTENER ////

    function addExpandAndCopyFeature() {
        // 1) Locate the table
        const table = document.querySelector('table.table-bordered.table-striped.table-condensed');
        if (!table) {
            console.error('Table not found!');
            return;
        }

        // 2) Find the <th> for "Model Description"
        let descriptionColumnIndex = -1;
        const headerCells = table.querySelectorAll('thead th');
        headerCells.forEach((th, idx) => {
            if (th.textContent.trim() === 'Model Description') {
                descriptionColumnIndex = idx;
            }
        });

        if (descriptionColumnIndex === -1) {
            console.warn('No "Model Description" column found in <thead><th>!');
            return;
        }

        // 3) Grab that header cell
        const descriptionHeaderCell = headerCells[descriptionColumnIndex];

        // Position it relative so our button can be placed absolute
        descriptionHeaderCell.style.position = 'relative';

        // 4) Create the expand/collapse button
        const expandBtn = document.createElement('button');
        expandBtn.textContent = '↔️';
        expandBtn.style.position = 'absolute';
        expandBtn.style.top = '-2%';
        expandBtn.style.transform = 'translateY(-50%)';
        expandBtn.style.left = '0px';
        expandBtn.style.zIndex = '9999';
        expandBtn.title = 'Expand Description';
        expandBtn.style.cursor = 'pointer';

        // Append to the header cell so the cell remains clickable for sorting
        descriptionHeaderCell.style.position = 'relative';  // ensure parent can position child absolutely
        descriptionHeaderCell.appendChild(expandBtn);


        // 5) Determine which cells are in the “Model Description” column
        // (Add 1 because nth-child is 1-based)
        const descriptionCells = table.querySelectorAll(`tbody tr td:nth-child(${descriptionColumnIndex + 1})`);

        // 6) Create a helper to copy text
        function copyToClipboard(text) {
            const textarea = document.createElement('textarea');
            textarea.value = text;
            document.body.appendChild(textarea);
            textarea.select();
            document.execCommand('');
            document.body.removeChild(textarea);
        }

        // 7) Expand/collapse toggle function for each cell
        function toggleCellDescription(cell, showFull) {
            const fullDescription = cell.dataset.fullDescription || cell.textContent.trim();
            const initialText = cell.dataset.initialText || fullDescription.split(' ')[0];

            // Save once if not stored yet
            if (!cell.dataset.fullDescription) {
                cell.dataset.fullDescription = fullDescription;
                cell.dataset.initialText = initialText;
            }

            // If showFull is true, show the full text; otherwise, truncate
            cell.textContent = showFull ? fullDescription : (initialText + '...');
        }

        // 8) Prepare all cells: store their data, truncate them by default,
        //    and attach a click-to-copy event
        descriptionCells.forEach(cell => {
            const fullDescription = cell.textContent.trim();
            const firstWord = fullDescription.split(' ')[0];
            cell.dataset.fullDescription = fullDescription;
            cell.dataset.initialText = firstWord;

            // Start truncated
            cell.textContent = firstWord + '...';

            // Make it look clickable
            cell.style.cursor = 'pointer';
            cell.title = 'Click to copy full description';

            // Clicking a cell copies the *full* text
            cell.addEventListener('click', () => {
                copyToClipboard(cell.dataset.fullDescription);

                // Brief feedback
                const originalText = cell.textContent;
                cell.textContent = 'Copied!✔️';
                setTimeout(() => {
                    cell.textContent = originalText;
                }, 700);
            });
        });

        // 9) Expand/Collapse all cells when the button is clicked
        let isExpanded = false;
        expandBtn.addEventListener('click', () => {
            isExpanded = !isExpanded;
            descriptionCells.forEach(cell => toggleCellDescription(cell, isExpanded));
            expandBtn.title = isExpanded ? 'Collapse Description' : 'Expand Description';
        });
    }

    // Execute
    addExpandAndCopyFeature();











    ///// ◀️📊 PULL BROKEN SERIAL FROM URL //////////////


    let containerVisible = false; // Variable to track container visibility

    // Function to toggle the visibility of the container and update button text
    const toggleContainerVisibility = () => {
        const container = document.getElementById('serialIdContainer');
        const toggleButton = document.getElementById('toggleButton'); // Get the toggle button element
        if (!containerVisible) {
            container.style.right = '0'; // Slide in from the right
            toggleButton.textContent = '❌'; // Change button text when container is visible
        } else {
            container.style.right = '-890px'; // Slide out to the right
            toggleButton.textContent = '🛠️'; // Change button text when container is hidden
        }
        containerVisible = !containerVisible; // Toggle visibility
    };

    // Create toggle button
    const toggleButton = document.createElement('button');
    toggleButton.id = 'toggleButton'; // Set button id
    toggleButton.textContent = '🛠️';
    toggleButton.onclick = toggleContainerVisibility;
    toggleButton.classList.add('toggleSerialsButton');
    toggleButton.setAttribute('title', '⚠️ Open Serial Repair History Data Table'); // Adding tooltip message
    document.body.appendChild(toggleButton);




    /////⛔ GET BROKEN PART DETAILS //////////

    const fetchBrokenSerialId = async (partId) => {
        const url = `https://mobility.AMAZON.com/part/part/rma/${partId}`;
        console.log('Fetching Broken Serial URL:', url);
        const response = await fetch(url);
        const html = await response.text();
        console.log('BROKEN PART Fetched HTML:', html);
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');

        // Construct XPath expressions for the latest and previous tables
        const latestTableXPath = '/html/body/div[3]/div/div[3]/div/div/table';
        const previousTableXPath = '/html/body/div[3]/div/div[2]/div/div/table'; // Updated previous table XPath

        // Evaluate the latest table XPath
        const latestTable = doc.evaluate(latestTableXPath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

        // Set the XPath for serial ID and model based on whether the latest table exists
        let serialIdXPath, modelXPath;
        if (latestTable) {
            serialIdXPath = `${latestTableXPath}/tbody/tr[1]/td[3]/table/tbody/tr[2]/td[2]`;
            modelXPath = `${latestTableXPath}/tbody/tr[1]/td[3]/table/tbody/tr[3]/td[2]`;
        } else {
            serialIdXPath = `${previousTableXPath}/tbody/tr[2]/td[3]/table/tbody/tr[2]/td[2]`; // Updated serial ID XPath for previous table
            modelXPath = `${previousTableXPath}/tbody/tr[2]/td[3]/table/tbody/tr[3]/td[2]`; // Updated model XPath for previous table
        }

        // Evaluate XPath expressions for serial ID and model
        const serialIdNode = doc.evaluate(serialIdXPath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        const modelNode = doc.evaluate(modelXPath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

        // Log and return serial ID
        if (serialIdNode) {
            console.log('Serial ID Found:', serialIdNode.textContent.trim());
            return serialIdNode.textContent.trim();
        } else {
            console.log('Serial ID Not Found');
            // Add backup XPath for serial ID
            const backupSerialIdXPath = '/html/body/div[3]/div/div[2]/div/div/table/tbody/tr/td[3]/table/tbody/tr[2]/td[2]';
            const backupSerialIdNode = doc.evaluate(backupSerialIdXPath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if (backupSerialIdNode) {
                console.log('Backup Serial ID Found:', backupSerialIdNode.textContent.trim());
                return backupSerialIdNode.textContent.trim();
            } else {
                console.log('Backup Serial ID Not Found');
                return 'NO REPAIR HISTORY';
            }
        }
    };

    const fetchBrokenModel = async (partId) => {
        const url = `https://mobility.AMAZON.com/part/part/rma/${partId}`;
        console.log('Fetching Broken Model URL:', url);
        const response = await fetch(url);
        const html = await response.text();
        console.log('Fetched Broken Model HTML:', html);
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');

        // Construct XPath expressions for the latest and previous tables
        const latestTableXPath = '/html/body/div[3]/div/div[3]/div/div/table';
        const previousTableXPath = '/html/body/div[3]/div/div[2]/div/div/table'; // Updated previous table XPath

        // Evaluate the latest table XPath
        const latestTable = doc.evaluate(latestTableXPath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

        // Set the XPath for the model based on whether the latest table exists
        let modelXPath;
        if (latestTable) {
            modelXPath = `${latestTableXPath}/tbody/tr[1]/td[3]/table/tbody/tr[3]/td[2]`;
        } else {
            modelXPath = `${previousTableXPath}/tbody/tr[2]/td[3]/table/tbody/tr[3]/td[2]`; // Updated model XPath for previous table
        }

        // Evaluate XPath expression for the model
        const modelNode = doc.evaluate(modelXPath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

        // Log and return model
        if (modelNode) {
            console.log('Broken Model Found:', modelNode.textContent.trim());
            return modelNode.textContent.trim();
        } else {
            console.log('Broken Model Not Found');
            // Add backup XPath for model
            const backupModelXPath = '/html/body/div[3]/div/div[2]/div/div/table/tbody/tr/td[3]/table/tbody/tr[3]/td[2]';
            const backupModelNode = doc.evaluate(backupModelXPath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if (backupModelNode) {
                console.log('Backup Model Found:', backupModelNode.textContent.trim());
                return backupModelNode.textContent.trim();
            } else {
                console.log('Backup Model Not Found');
                return 'NO REPAIR HISTORY';
            }
        }
    };






    /////🔄 GET CONSUMED PART DETAILS //////////

    const fetchConsumedSerialId = async (partId) => {
        const url = `https://mobility.AMAZON.com/part/part/rma/${partId}`;
        console.log('Fetching Consumed Serial URL:', url);
        const response = await fetch(url);
        const html = await response.text();
        console.log('Fetched Consumed Serial HTML:', html);
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');

        // Construct XPath expressions for the latest and previous tables
        const latestTableXPath = '/html/body/div[3]/div/div[3]/div/div/table';
        const previousTableXPath = '/html/body/div[3]/div/div[2]/div/div/table'; // Updated previous table XPath

        // Evaluate the latest table XPath
        const latestTable = doc.evaluate(latestTableXPath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

        // Set the XPath for serial ID based on whether the latest table exists
        let serialIdXPath;
        if (latestTable) {
            serialIdXPath = `${latestTableXPath}/tbody/tr[1]/td[4]/table/tbody/tr[2]/td[2]`;
        } else {
            serialIdXPath = `${previousTableXPath}/tbody/tr[2]/td[4]/table/tbody/tr[2]/td[2]`; // Updated serial ID XPath for previous table
        }

        // Evaluate XPath expression for serial ID
        const serialIdNode = doc.evaluate(serialIdXPath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

        // Log and return serial ID
        if (serialIdNode) {
            console.log('Consumed Serial ID Found:', serialIdNode.textContent.trim());
            return serialIdNode.textContent.trim();
        } else {
            console.log('Consumed Serial ID Not Found');
            // Add backup XPath for consumed serial ID
            const backupSerialIdXPath = '/html/body/div[3]/div/div[2]/div/div/table/tbody/tr/td[4]/table/tbody/tr[2]/td[2]';
            const backupSerialIdNode = doc.evaluate(backupSerialIdXPath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if (backupSerialIdNode) {
                console.log('Backup Consumed Serial ID Found:', backupSerialIdNode.textContent.trim());
                return backupSerialIdNode.textContent.trim();
            } else {
                console.log('Backup Consumed Serial ID Not Found');
                return 'Serial Not Found';
            }
        }
    };

    const fetchConsumedModel = async (partId) => {
        const url = `https://mobility.AMAZON.com/part/part/rma/${partId}`;
        console.log('Fetching Consumed Model URL:', url);
        const response = await fetch(url);
        const html = await response.text();
        console.log('Fetched Consumed Model HTML:', html);
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');

        // Construct XPath expressions for the latest and previous tables
        const latestTableXPath = '/html/body/div[3]/div/div[3]/div/div/table';
        const previousTableXPath = '/html/body/div[3]/div/div[2]/div/div/table'; // Updated previous table XPath

        // Evaluate the latest table XPath
        const latestTable = doc.evaluate(latestTableXPath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

        // Set the XPath for the model based on whether the latest table exists
        let modelXPath;
        if (latestTable) {
            modelXPath = `${latestTableXPath}/tbody/tr[1]/td[4]/table/tbody/tr[3]/td[2]`;
        } else {
            modelXPath = `${previousTableXPath}/tbody/tr[2]/td[4]/table/tbody/tr[3]/td[2]`; // Updated model XPath for previous table
        }

        // Evaluate XPath expression for the model
        const modelNode = doc.evaluate(modelXPath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

        // Log and return model
        if (modelNode) {
            console.log('Consumed Model Found:', modelNode.textContent.trim());
            return modelNode.textContent.trim();
        } else {
            console.log('Consumed Model Not Found');
            // Add backup XPath for consumed model
            const backupModelXPath = '/html/body/div[3]/div/div[2]/div/div/table/tbody/tr/td[4]/table/tbody/tr[3]/td[2]';
            const backupModelNode = doc.evaluate(backupModelXPath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if (backupModelNode) {
                console.log('Backup Consumed Model Found:', backupModelNode.textContent.trim());
                return backupModelNode.textContent.trim();
            } else {
                console.log('Backup Consumed Model Not Found');
                return 'Model not found';
            }
        }
    };







    // Function to check if a serial number matches any serial numbers in the main table
    const serialMatchesMainTable = (serialNumber) => {
        const mainTableSerials = Array.from(document.querySelectorAll('table.table-bordered.table-striped.table-condensed tbody tr td:nth-child(3)'));
        return mainTableSerials.some(cell => cell.textContent.trim().replace(/\s*\(\s*RMA\s*\)$/, '') === serialNumber);
    };

    // Array to hold broken serials
    const brokenSerials = [];

    // Array to hold consumed serials
    const consumedSerials = [];



    // Function to display serial IDs
    const displaySerialIds = async () => {
        const tableRows = document.querySelectorAll('table.table-bordered.table-striped.table-condensed tbody tr');
        let rowCount = 0; // Initialize count of rows added
        const totalCount = tableRows.length; // Total number of rows in the main table

        // Create container for serial IDs
        const container = document.createElement('div');
        container.id = 'serialIdContainer';
        container.style.position = 'fixed';
        container.style.background = '#aeb6bf';
        container.style.width = '800px';
        container.style.top = '20px';
        container.style.marginTop = '56px';
        container.style.right = '-890px'; // Initially hide off the screen from the top right
        container.style.padding = '10px';
        container.style.border = '1px solid #000'; // Add border
        container.style.borderRadius = '5px';
        container.style.zIndex = '9999';
        container.style.overflowY = 'auto'; // Add scrollbar for overflow
        container.style.maxHeight = '80vh'; // Set maximum height to 80% of viewport height
        container.style.transition = 'opacity 0.3s ease, right 0.3s ease'; // Add transition for smooth sliding motion
        container.style.border = '4px solid #ccc'; // Lighter border color
        container.style.borderTop = '4px solid #f5f5f5'; // Lighter color on top for shine effect
        container.style.borderLeft = '4px solid #f5f5f5'; // Lighter color on left for shine effect
        container.style.backgroundClip = 'padding-box'; // Apply background within padding
        container.style.fontFamily = 'Arial, sans-serif'; // Set font family with fallbacks
        container.style.fontWeight = 'bold'; // Bold text
        container.style.textShadow = '1px 1px 2px rgba(255, 255, 255, 0.5)'; // Soft text shadow for 3D effect

        // Helper function to create cell elements with specified text content
        const createCell = (textContent) => {
            const cell = document.createElement('div');
            cell.textContent = textContent;
            cell.style.flex = '1';
            cell.style.textAlign = 'left';
            cell.style.fontFamily = 'Arial, sans-serif'; // Change font family to Arial or sans-serif for a crisp and professional look
            cell.style.fontWeight = 'normal'; // Make text bold
            return cell;
        };

        // Helper function to create heading elements
        const createHeading = (text, symbol) => {
            const heading = createCell(`${symbol} ${text}`);
            heading.style.fontSize = '16px';
            return heading;
        };

        // Helper function to create copy button
        const createCopyButton = (onClick) => {
            const button = document.createElement('div');
            button.textContent = '📋';
            button.style.fontSize = '15px';
            button.style.cursor = 'pointer';
            button.style.marginRight = '5px'; // Adjust margin-right here
            button.style.marginLeft = '-135px'; // Adjust margin-right here
            button.onclick = onClick;
            return button;
        };

        // Helper function to create search button
        const createSearchButton = (onClick) => {
            const button = document.createElement('div');
            button.textContent = '🔗';
            button.style.fontSize = '15px';
            button.style.cursor = 'pointer';
            button.style.marginRight = '55px'; // Adjust margin-right here
            button.onclick = onClick;
            return button;
        };

        // Append header elements to container
        const headingRow = document.createElement('div');
        headingRow.style.display = 'flex';
        headingRow.style.alignItems = 'left'; // Align items vertically
        headingRow.style.fontWeight = 'bold';
        headingRow.style.borderBottom = '1px solid #000'; // Add border
        headingRow.style.paddingBottom = '5px'; // Add padding
        headingRow.style.marginBottom = '5px';
        headingRow.style.position = 'sticky'; // Make the header sticky
        headingRow.style.top = '0'; // Stick it to the top of the viewport
        headingRow.style.backgroundColor = '#aeb6bf'; // Set background color


        //// COPY SNs / MOBILITY URL BUTTONS ///////



        const brokenSNHeading = createHeading('SERIAL', '⛔');
        const brokenCopyButton = createCopyButton(() => {
            const validBrokenSerials = brokenSerials.filter(serial => serial !== 'NO REPAIR HISTORY');
            const brokenSerialsText = validBrokenSerials.join('\n');
            navigator.clipboard.writeText(brokenSerialsText)
                .then(() => {
                showSuccessTick(brokenCopyButton);
            })
                .catch((error) => {
                console.error('Failed to copy broken serials: ', error);
            });
        });
        const brokenSearchButton = createSearchButton(() => {
            const validBrokenSerials = brokenSerials.filter(serial => serial !== 'NO REPAIR HISTORY');
            const url = 'https://mobility.AMAZON.com/part/search?search_type=all&search_string=' + validBrokenSerials.join('+') + '&max_rows=6000&query=GO';
            window.open(url, '_blank');
        });

        const brokenModelHeading = createHeading('MODEL', '⛔');
        brokenModelHeading.style.marginLeft = '-10px';

        const consumedSNHeading = createHeading('SERIAL', '🔄');
        const consumedCopyButton = createCopyButton(() => {
            const validConsumedSerials = consumedSerials.filter(serial => serial !== 'Serial Not Found');
            const consumedSerialsText = validConsumedSerials.join('\n');
            navigator.clipboard.writeText(consumedSerialsText)
                .then(() => {
                showSuccessTick(consumedCopyButton);
            })
                .catch((error) => {
                console.error('Failed to copy consumed serials: ', error);
            });
        });
        const consumedSearchButton = createSearchButton(() => {
            const validConsumedSerials = consumedSerials.filter(serial => serial !== 'Serial Not Found');
            const url = 'https://mobility.AMAZON.com/part/search?search_type=all&search_string=' + validConsumedSerials.join('+') + '&max_rows=6000&query=GO';
            window.open(url, '_blank');
        });

        const consumedModelHeading = createHeading('MODEL', '🔄');
        consumedModelHeading.style.marginRight = '-40px';



        // Create element for displaying count
        const countElement = document.createElement('div');
        countElement.textContent = '0/0'; // Initial count
        countElement.style.fontSize = '14px';
        countElement.style.textAlign = 'right';
        countElement.style.marginTop = '5px';

        headingRow.appendChild(brokenSNHeading);
        headingRow.appendChild(brokenCopyButton);
        headingRow.appendChild(brokenSearchButton);
        headingRow.appendChild(brokenModelHeading);
        headingRow.appendChild(consumedSNHeading);
        headingRow.appendChild(consumedCopyButton);
        headingRow.appendChild(consumedSearchButton);
        headingRow.appendChild(consumedModelHeading);
        headingRow.appendChild(countElement);
        container.appendChild(headingRow);


        document.body.appendChild(container);

        // Function to temporarily change the button content to a green tick emoji for 1 second
        const showSuccessTick = (button) => {
            button.textContent = '✔️'; // Change button content to green tick emoji
            setTimeout(() => {
                button.textContent = '📋'; // Change button content back to copy symbol after 1 second
            }, 1000);
        };

        // Iterate through table rows
        for (const row of tableRows) {
            const partId = row.querySelectorAll('td')[2 - 1].textContent.trim();
            const brokenSerialId = await fetchBrokenSerialId(partId);
            const consumedSerialId = await fetchConsumedSerialId(partId);
            const brokenModel = await fetchBrokenModel(partId);
            const consumedModel = await fetchConsumedModel(partId);

            // Create row elements
            const serialRow = document.createElement('div');
            serialRow.style.display = 'flex';
            serialRow.style.justifyContent = 'space-between';
            serialRow.style.marginBottom = '5px';

            // Create numbering element
            const numberingElement = document.createElement('div');
            numberingElement.textContent = `${rowCount + 1}.`;
            numberingElement.style.marginRight = '5px'; // Adjust spacing here

            // Create broken and consumed cells
            const brokenSNCell = createCell(brokenSerialId);

            const brokenModelCell = createCell(brokenModel);
            brokenModelCell.style.marginLeft = '-25px';
            brokenModelCell.style.marginRight = '30px'

            const consumedSNCell = createCell(consumedSerialId);

            const consumedModelCell = createCell(consumedModel);
            consumedModelCell.style.marginLeft = '-25px';

            // Highlight broken serial if it matches main table
            if (serialMatchesMainTable(brokenSerialId)) {
                brokenSNCell.innerHTML = `<span style="background-color: lightcoral; padding: 3px 5px; border: 1px solid #ccc; border-radius: 3px; box-shadow: 0 0 5px rgba(0, 0, 0, 0.3) inset;">${brokenSerialId}</span>`;
                brokenSNCell.style.fontWeight = 'bold';
            }

            // Highlight consumed serial if it matches main table
            if (serialMatchesMainTable(consumedSerialId)) {
                consumedSNCell.innerHTML = `<span style="background-color:  #2ecc71; padding: 3px 5px; border: 1px solid #ccc; border-radius: 3px; box-shadow: 0 0 5px rgba(0, 0, 0, 0.3) inset;">${consumedSerialId}</span>`;
                consumedSNCell.style.fontWeight = 'bold';
            }

            // Append cells to row
            serialRow.appendChild(numberingElement); // Append numbering element
            serialRow.appendChild(brokenSNCell);
            serialRow.appendChild(brokenModelCell);
            serialRow.appendChild(consumedSNCell);
            serialRow.appendChild(consumedModelCell);
            container.appendChild(serialRow);

            // Update row count and display it in the container
            rowCount++;
            const countText = `${rowCount}/${totalCount}`;
            countElement.textContent = countText;

            // Store broken and consumed serials
            brokenSerials.push(brokenSerialId);
            consumedSerials.push(consumedSerialId);

            // Log models for debugging
            console.log("Broken Model:", brokenModel);
            console.log("Consumed Model:", consumedModel);
            console.log("Models Match:", brokenModel === consumedModel);

            // Apply orange color to model cells if they match immediately after creation
            if (brokenModel === consumedModel) {
                brokenModelCell.innerHTML = `<span style="background-color: orange; padding: 3px 5px; border: 1px solid #ccc; border-radius: 3px; box-shadow: 0 0 5px rgba(0, 0, 0, 0.3) inset;">${brokenModel}</span>`;
                consumedModelCell.innerHTML = `<span style="background-color: orange; padding: 3px 5px; border: 1px solid #ccc; border-radius: 3px; box-shadow: 0 0 5px rgba(0, 0, 0, 0.3) inset;">${consumedModel}</span>`;
            }
        }
    };

    // Call the function to display serial IDs
    displaySerialIds();




    ///🏁 FETCH ALL LAST USERS //////



    let UserContainerVisible = false; // Variable to track container visibility
    let popupVisible = false; // Variable to track popup visibility
    let currentPopup = null; // Reference to the current popup

    // Function to toggle the visibility of the last updated users container
    const toggleLastUserContainerVisibility = () => {
        const container = document.getElementById('lastUpdatedContainer');
        const toggleLastUserButton = document.getElementById('toggleLastUpdated'); // Get the toggle button element
        if (container) {
            if (!UserContainerVisible) {
                container.style.right = '0'; // Slide in from the right
                toggleLastUserButton.textContent = '❌'; // Change button text when container is visible
            } else {
                container.style.right = '-630px'; // Slide out to the right
                toggleLastUserButton.textContent = '🙎‍♂️'; // Change button text when container is hidden
            }
            UserContainerVisible = !UserContainerVisible; // Toggle visibility
        } else {
            fetchAllLastUpdatedUsersAndActions(); // If container doesn't exist, fetch and display it
        }
    };

    // Initialize the toggle button to fetch and display last updated users
    const toggleLastUserButton = document.createElement('button');
    toggleLastUserButton.id = 'toggleLastUpdated';
    toggleLastUserButton.textContent = '🙎‍♂️';
    toggleLastUserButton.onclick = toggleLastUserContainerVisibility;
    toggleLastUserButton.classList.add('toggleLastUsers');
    toggleLastUserButton.setAttribute('title', '⚠️ Open Last Users Data Table'); // Adding tooltip message
    document.body.appendChild(toggleLastUserButton);

    // Function to fetch last updated users and actions for a specific partId
    const fetchLastUpdatedUserAndAction = async (partId) => {
        try {
            const url = `https://mobility.AMAZON.com/part/part/${partId}`;
            console.log('Accessing URL LastUpdatedUser:', url);

            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Content-Type': 'text/html',
                },
            });

            const htmlText = await response.text();
            console.log('LastUserAction HTML response received:', htmlText);

            // Create a temporary div to parse the HTML response
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = htmlText;

            console.log('Parsed HTML LastUpdatedUser:', tempDiv);

            // Find the sort table
            const sortTable = tempDiv.querySelector('#sortTable');

            if (!sortTable) {
                console.log('Table with ID "sortTable" not found.');
                return;
            }

            console.log('Table with ID "sortTable" found:', sortTable);

            // Get the last row of the table
            const cells = sortTable.querySelectorAll('tbody tr:last-child td');
            const userInfo = cells[3].textContent.trim();
            const matches = userInfo.match(/\(([^)]+)\)/);

            let lastUpdatedUsername = 'Unknown';
            let actionData = 'No Action'; // Default action data if not found
            let actionDate = 'No Date'; // Default action date if not found

            if (matches && matches.length > 1) {
                lastUpdatedUsername = matches[1];
                console.log('Last updated username:', lastUpdatedUsername);

                // Fetch action data from cells 1 and 2 (two cells before the user info cell)
                const actionCell1 = cells[1]; // Cell 1 for action data
                const actionCell2 = cells[2]; // Cell 2 for action data
                actionData = `${actionCell1.textContent.trim()}  |  ${actionCell2.textContent.trim()}`;
                console.log('Action data:', actionData);

                // Fetch action date from cell 0 (three cells before the user info cell)
                const actionDateCell = cells[0]; // Cell 0 for action date
                actionDate = actionDateCell.textContent.trim();
                console.log('Action date:', actionDate);
            } else {
                console.log('Username format incorrect.');
            }

            return { lastUpdatedUsername, actionData, actionDate, htmlText: sortTable.outerHTML };
        } catch (error) {
            console.error('Error fetching or processing data:', error);
            return { lastUpdatedUsername: 'Unknown', actionData: 'Unknown', actionDate: 'Unknown', htmlText: '' };
        }
    };

    const fetchAllLastUpdatedUsersAndActions = async () => {
        try {
            // Check if the container already exists, if yes, remove it
            const existingContainer = document.getElementById('lastUpdatedContainer');
            if (existingContainer) {
                existingContainer.remove();
            }

            const rows = document.querySelectorAll('table.table-bordered.table-striped.table-condensed tbody tr');
            const container = document.createElement('div');
            container.id = 'lastUpdatedContainer';
            container.style.position = 'fixed';
            container.style.background = ' #85c1e9 ';
            container.style.width = '600px'; // Adjusted width to accommodate the new column
            container.style.top = '20px';
            container.style.right = '-630px'; // Initially hide off the screen from the top right
            container.style.padding = '10px';
            container.style.border = '1px solid #000'; // Add border
            container.style.borderRadius = '5px';
            container.style.zIndex = '9999';
            container.style.overflowY = 'auto'; // Add scrollbar for overflow
            container.style.maxHeight = '80vh'; // Set maximum height to 80% of viewport height
            container.style.transition = 'opacity 0.3s ease, right 0.3s ease'; // Add transition for smooth sliding motion
            container.style.border = '4px solid #ccc'; // Lighter border color
            container.style.borderTop = '4px solid #f5f5f5'; // Lighter color on top for shine effect
            container.style.borderLeft = '4px solid #f5f5f5'; // Lighter color on left for shine effect
            container.style.backgroundClip = 'padding-box'; // Apply background within padding
            container.style.fontFamily = 'Arial, sans-serif'; // Set font family with fallbacks
            container.style.fontWeight = 'bold'; // Bold text
            container.style.textShadow = '1px 1px 2px rgba(255, 255, 255, 0.5)'; // Soft text shadow for 3D effect
            document.body.appendChild(container);

            // Add CSV download button
            const downloadButtonContainer = document.createElement('div');
            downloadButtonContainer.style.display = 'flex';
            downloadButtonContainer.style.alignItems = 'center'; // Center align items vertically
            downloadButtonContainer.style.justifyContent = 'center'; // Center align items horizontally
            downloadButtonContainer.style.marginBottom = '10px'; // Add bottom margin for spacing
            downloadButtonContainer.style.marginRight = '20px';
            downloadButtonContainer.style.padding = '0px';

            const downloadButton = document.createElement('button');
            downloadButton.textContent = '💾';
            downloadButton.marginRight = '10px';
            downloadButton.addEventListener('click', () => {
                downloadCSV();
            });
            downloadButtonContainer.appendChild(downloadButton);

            // Add title after download button
            const titleContainer = document.createElement('div');
            titleContainer.style.display = 'flex';
            titleContainer.style.justifyContent = 'space-between';
            titleContainer.style.alignItems = 'center';
            titleContainer.style.marginLeft = '8px';

            container.appendChild(downloadButtonContainer);

            const title = document.createElement('div');
            title.textContent = 'Last Interaction Data';
            title.style.fontWeight = 'bold';
            title.style.fontSize = '18px';
            titleContainer.appendChild(title);

            // Create a div for row count
            const countElement = document.createElement('div');
            countElement.style.fontWeight = 'bold';
            countElement.style.fontSize = '18px';
            countElement.style.color = 'black'; // Black text color
            countElement.style.backgroundColor = '#85c1e9'; // Background color matching container
            countElement.style.padding = '5px 10px'; // Padding for the count
            countElement.style.borderRadius = '5px'; // Rounded corners
            countElement.style.marginLeft = '10px'; // Left margin for spacing

            let rowCount = 0; // Initialize row count
            const totalCount = rows.length; // Total number of rows

            const countText = `${rowCount}/${totalCount}`;
            countElement.textContent = countText;
            titleContainer.appendChild(countElement);

            downloadButtonContainer.appendChild(titleContainer);

            // Add table to display last updated users and actions
            const table = document.createElement('table');
            table.id = 'lastUpdatedTable';
            table.style.width = '100%';
            container.appendChild(table);

            const thead = document.createElement('thead');
            const tbody = document.createElement('tbody');
            table.appendChild(thead);
            table.appendChild(tbody);

            // Add table headers
            const headers = ['SERIALS', 'USER 👤', 'LAST CHANGE 🔄', '📅🕒']; // Added "DATE" header

            // Create header row
            const headerRow = document.createElement('tr');
            headerRow.style.fontWeight = 'bold';
            headerRow.style.paddingBottom = '5px'; // Add padding
            headerRow.style.marginBottom = '10px';
            headerRow.style.fontSize = '14px';
            headerRow.style.backgroundColor = '#85c1e9'; // Set background color

            headers.forEach((headerText, index) => {
                // Create individual header cell
                const th = document.createElement('th');
                th.textContent = headerText;
                th.style.padding = '4px'; // Add padding to header cell
                th.style.marginRight = '20px'; // Add right margin to match cell margin
                th.style.textAlign = 'center'; // Center align header text

                headerRow.appendChild(th);
            });

            thead.appendChild(headerRow); // Append the header row to the table's header

            // Fetch data and append to the table for each row
            for (let i = 0; i < rows.length; i++) {
                const cells = rows[i].querySelectorAll('td');
                const serialNumber = cells[2].textContent.trim().replace(/\s*\(\s*RMA\s*\)$/, ''); // Trim and remove RMA
                const partId = cells[1].textContent.trim().replace(/\s*📦$/, ''); // Trim and remove emoji

                // Fetch last updated user, action, and date for the current partId
                const { lastUpdatedUsername, actionData, actionDate, htmlText } = await fetchLastUpdatedUserAndAction(partId);

                // Split the action data into two parts
                const [action1, action2] = actionData.split('  |  ');

                // Replace action text with emojis based on conditions
                const actionEmoji1 = action1 === 'Tag Added' ? '🏷️' : action1 === 'Bin' ? '🗑️' : action1 === 'State' ? '🔄' : action1;
                const actionEmoji2 = action2 === 'PENDING_SANITIZATION_IDENTIFICATION' ? 'PENDING SAN ID' : action2 === 'Bin' ? '🗑️' : action2;

                // Format the action date
                const [day, month, year] = actionDate.split('-');
                const formattedDate = `${day}-${month}-${year}`;

                // Create row for the last updated user and action, and append to the table
                const row = document.createElement('tr');

                const serialNumberCell = document.createElement('td');
                serialNumberCell.textContent = serialNumber; // Use serial number from the row
                serialNumberCell.style.padding = '2px'; // Add padding to the cell
                serialNumberCell.style.textAlign = 'left'; // Left-align cell content
                serialNumberCell.style.border = '1px solid #000'; // Add border

                // Add click event listener to toggle the popup
                serialNumberCell.addEventListener('click', async (event) => {
                    event.stopPropagation(); // Prevent click event from bubbling up
                    const newHtmlText = await fetchLastUpdatedUserAndAction(partId).then(res => res.htmlText);
                    togglePopup(newHtmlText, serialNumberCell);
                });

                row.appendChild(serialNumberCell);

                const usernameCell = document.createElement('td');
                usernameCell.textContent = lastUpdatedUsername;
                usernameCell.style.padding = '2px'; // Add padding to the cell
                usernameCell.style.backgroundColor = '#fff'; // Set background color to white
                usernameCell.style.border = '1px solid #000'; // Add border
                usernameCell.style.borderRadius = '5px'; // Add border radius for curved corners
                usernameCell.style.textAlign = 'center'; // Center-align cell content
                row.appendChild(usernameCell);

                const actionCell = document.createElement('td');
                actionCell.style.padding = '2px'; // Add padding to the cell
                actionCell.style.backgroundColor = 'orange'; // Set background color to orange
                actionCell.style.border = '1px solid #000'; // Add border
                actionCell.style.borderRadius = '5px'; // Add border radius for curved corners
                actionCell.style.textAlign = 'left'; // Center-align cell content

                // Create spans to display the two actions with separation
                const actionSpan1 = document.createElement('span');
                actionSpan1.textContent = actionEmoji1; // Use emoji instead of action text
                const actionSpan2 = document.createElement('span');
                actionSpan2.textContent = actionEmoji2; // Use emoji instead of action text

                // Append action spans to the action cell with separator
                actionCell.appendChild(actionSpan1);
                actionCell.appendChild(document.createTextNode(' | '));
                actionCell.appendChild(actionSpan2);

                row.appendChild(actionCell);

                const dateCell = document.createElement('td');
                dateCell.textContent = formattedDate; // Use formatted date
                dateCell.style.padding = '2px'; // Add padding to the cell
                dateCell.style.backgroundColor = '#fff'; // Set background color to white
                dateCell.style.border = '1px solid #000'; // Add border
                dateCell.style.borderRadius = '5px'; // Add border radius for curved corners
                dateCell.style.textAlign = 'center'; // Center-align cell content
                row.appendChild(dateCell);

                tbody.appendChild(row); // Append the row to the table's body

                // Increment the row count
                rowCount++;
                const countText = `${rowCount}/${totalCount}`;
                countElement.textContent = countText; // Update the row count text
            }

            // Append the table body to the table
            table.appendChild(tbody);

        } catch (error) {
            console.error('Error fetching or processing data:', error);
        }
    };

    // Function to show or hide the popup with the HTML content
    const togglePopup = (htmlText, targetElement) => {
        if (popupVisible) {
            hidePopup();
        }
        showPopup(htmlText, targetElement);
    };

    // Function to show the popup with the HTML content
    const showPopup = (htmlText, targetElement) => {
        // If there is an existing popup, remove it first
        if (currentPopup) {
            currentPopup.remove();
            currentPopup = null;
        }

        const popup = document.createElement('div');
        popup.id = 'htmlPopup';
        popup.style.position = 'fixed'; // Use fixed position for consistent placement
        popup.style.background = 'white';
        popup.style.border = '1px solid black';
        popup.style.padding = '10px';
        popup.style.zIndex = '10000';
        popup.style.width = '1000px'; // Increased width
        popup.style.maxHeight = '300px';
        popup.style.overflowY = 'auto';
        popup.style.left = '20px'; // Fixed position from the left side of the screen
        popup.style.top = '50%'; // Center vertically
        popup.style.transform = 'translateY(-50%)'; // Adjust for perfect vertical centering

        // Create a title for the popup
        const title = document.createElement('h3');
        title.textContent = `Audit History of SN: ${targetElement.textContent}`;
        title.style.textAlign = 'center';
        title.style.fontSize = '16px';
        title.style.color = 'black';
        title.style.borderRadius = '10px';
        title.style.marginBottom = '8px';
        title.style.backgroundColor = '#f1c40f';
        popup.appendChild(title);

        // Create a container div to apply styles to the table
        const tableContainer = document.createElement('div');
        tableContainer.innerHTML = htmlText;
        const table = tableContainer.querySelector('#sortTable');

        // Apply target="_blank" to all hyperlinks
        const links = table.querySelectorAll('a');
        links.forEach(link => {
            link.setAttribute('target', '_blank');
        });

        // Apply styles to the table
        table.style.width = '100%';
        table.style.borderCollapse = 'collapse';

        // Apply styles to the table headers
        const thElements = table.querySelectorAll('th');
        thElements.forEach(th => {
            th.style.padding = '5px';
            th.style.border = '1px solid black';
            th.style.textAlign = 'center';
            th.style.position = 'sticky'; // Make header sticky
            th.style.top = '0'; // Position at the top
            th.style.backgroundColor = '#34495e'; // Background color for headers
            th.style.zIndex = '1'; // Ensure headers stay on top
            th.style.color = 'white'; // Ensure text color is visible
        });

        // Apply styles to the table cells and add emojis before specific text
        const tdElements = table.querySelectorAll('td');
        tdElements.forEach(td => {
            td.style.padding = '5px';
            td.style.border = '1px solid black';
            td.style.textAlign = 'center';
            td.style.whiteSpace = 'nowrap'; // Prevent text wrapping

            // Add emojis based on specific keywords (case insensitive) and left-align text in this column
            if (/bin/i.test(td.textContent)) {
                td.style.textAlign = 'left'; // Left-align text
                td.innerHTML = '🗑️ ' + td.textContent; // Add bin emoji before the text
            } else if (/tag added/i.test(td.textContent)) {
                td.style.textAlign = 'left'; // Left-align text
                td.innerHTML = '🏷️ ' + td.textContent; // Add tag emoji before the text
            } else if (/user who has custody/i.test(td.textContent)) {
                td.style.textAlign = 'left'; // Left-align text
                td.innerHTML = '👤 ' + td.textContent; // Add user emoji before the text
            } else if (/state/i.test(td.textContent)) {
                td.style.textAlign = 'left'; // Left-align text
                td.innerHTML = '📝 ' + td.textContent; // Add state emoji before the text
            } else if (/tracking id/i.test(td.textContent)) {
                td.style.textAlign = 'left'; // Left-align text
                td.innerHTML = '🚚 ' + td.textContent; // Add tracking ID emoji before the text
            } else if (/cluster/i.test(td.textContent)) {
                td.style.textAlign = 'left'; // Left-align text
                td.innerHTML = '🖧 ' + td.textContent; // Add cluster emoji before the text
            } else if (/received at/i.test(td.textContent)) {
                td.style.textAlign = 'left'; // Left-align text
                td.innerHTML = '✍🏻 ' + td.textContent; // Add received at emoji before the text
            } else if (/model/i.test(td.textContent)) {
                td.style.textAlign = 'left'; // Left-align text
                td.innerHTML = '🔡 ' + td.textContent; // Add model emoji before the text
            } else if (/serial id/i.test(td.textContent)) {
                td.style.textAlign = 'left'; // Left-align text
                td.innerHTML = '🆔 ' + td.textContent; // Add serial ID emoji before the text
            } else if (/category/i.test(td.textContent)) {
                td.style.textAlign = 'left'; // Left-align text
                td.innerHTML = '🗂️ ' + td.textContent; // Add serial ID emoji before the text
            } else if (/tote asset id/i.test(td.textContent)) {
                td.style.textAlign = 'left'; // Left-align text
                td.innerHTML = '📦 ' + td.textContent; // Add serial ID emoji before the text
            }

            // Hyperlink the text in the "Modified By" column
            if (/modified by/i.test(td.previousElementSibling?.textContent)) {
                const matches = td.textContent.match(/\(([^)]+)\)/);
                if (matches && matches[1]) {
                    const user = matches[1];
                    const link = document.createElement('a');
                    link.href = `https://phonetool.EXAMPLE.com/users/${user}`;
                    link.target = '_blank';
                    link.textContent = `(${user})`;
                    td.innerHTML = td.textContent.replace(`(${user})`, '');
                    td.appendChild(link);
                }
            }
        });

        // Append the table to the popup
        popup.appendChild(table);
        document.body.appendChild(popup);

        popupVisible = true;
        currentPopup = popup;

        // Add event listener to close the popup when clicking outside
        document.addEventListener('click', handleClickOutside);
    };

    // Function to hide the popup
    const hidePopup = () => {
        if (currentPopup) {
            currentPopup.remove();
            currentPopup = null;
            popupVisible = false;
            document.removeEventListener('click', handleClickOutside);
        }
    };

    // Function to handle click outside the popup
    const handleClickOutside = (event) => {
        if (currentPopup && !currentPopup.contains(event.target)) {
            hidePopup();
        }
    };

    // Function to download CSV
    const downloadCSV = () => {
        const rows = document.querySelectorAll('#lastUpdatedTable tbody tr');

        let csvContent = 'SERIALS,USERNAME,ACTION,DATE\n';

        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            const serialNumber = cells[0].textContent;
            const username = cells[1].textContent;
            const action = cells[2].textContent;
            const date = cells[3].textContent;

            csvContent += `${serialNumber},${username},${action},${date}\n`;
        });

        const blob = new Blob([csvContent], { type: 'text/csv' });
        const url = window.URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = 'last_updated_data.csv';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    };

    // Call the function to display serial IDs, usernames, actions, and dates
    fetchAllLastUpdatedUsersAndActions();






    ////🧾 PART SUMMARY POP-UP //////// 


    (function() {
        // Function to find the index of the "Serial Id" column
        function findSerialIdColumnIndex() {
            const headerRow = document.querySelector('table thead tr');
            if (!headerRow) return -1;

            const headers = headerRow.querySelectorAll('th');
            for (let i = 0; i < headers.length; i++) {
                if (headers[i].textContent.trim().toLowerCase() === 'serial id') {
                    return i + 1; // CSS nth-child is 1-indexed
                }
            }
            return -1; // Column not found
        }

        // Function to handle hover and click on cells containing serial numbers
        function handleSerialNumberCellEvents() {
            const serialIdColumnIndex = findSerialIdColumnIndex();
            if (serialIdColumnIndex === -1) {
                console.error('Serial Id column not found');
                return;
            }

            const serialNumberCells = document.querySelectorAll(`table tbody tr td:nth-child(${serialIdColumnIndex})`);
            let activeCell = null; // Keep track of the active cell that triggered the pop-up

            serialNumberCells.forEach(cell => {
                // Show tooltip on hover
                cell.addEventListener('mouseenter', function(event) {
                    const tooltip = document.createElement('div');
                    tooltip.textContent = 'Click for part summary';
                    tooltip.style.position = 'absolute';
                    tooltip.style.backgroundColor = 'rgba(0, 0, 0, 0.7)';
                    tooltip.style.color = 'white';
                    tooltip.style.padding = '5px';
                    tooltip.style.borderRadius = '5px';
                    tooltip.style.zIndex = '9999';

                    // Position the tooltip to the right of the hovered cell
                    const cellRect = cell.getBoundingClientRect();
                    tooltip.style.top = (window.scrollY + cellRect.top) + 'px';
                    tooltip.style.left = (cellRect.right + 10) + 'px'; // Adjust the distance from the cell here

                    // Append the tooltip to the document body
                    document.body.appendChild(tooltip);

                    // Remove the tooltip when the mouse leaves the cell
                    cell.addEventListener('mouseleave', function() {
                        tooltip.remove();
                    });
                });

                // Show or hide part summary form on click
                cell.addEventListener('click', function(event) {
                    // Check if the clicked cell is the same as the one that triggered the current pop-up
                    if (activeCell === cell) {
                        // If it is, remove the pop-up form and reset activeCell
                        activeCell = null;
                        const popupForm = document.querySelector('.popup-form');
                        if (popupForm) {
                            popupForm.remove();
                        }
                    } else {
                        // If it's not, remove any existing pop-up form
                        const existingPopupForm = document.querySelector('.popup-form');
                        if (existingPopupForm) {
                            existingPopupForm.remove();
                        }

                        // Define flag mappings for different substrings
                        const flagMappings = {
                            'BH': 'Bahrain 🇧🇭', // Flag emoji for Bahrain
                            'US': 'USA 🇺🇸', // Flag emoji for USA
                            'ZA': 'South Africa 🇿🇦', // Flag emoji for South Africa
                            'AE': 'UAE 🇦🇪', // Flag emoji for UAE
                            'IT': 'Italy 🇮🇹', // Flag emoji for Italy
                            'DE': 'Germany 🇩🇪', // Flag emoji for Germany
                            'SW': 'Sweden 🇸🇪', // Flag emoji for Sweden
                            'LHR': 'UK 🇬🇧', // Flag emoji for UK
                            'ZHY': 'China 🇨🇳', // Flag emoji for China
                            'EU': 'Ireland 🇮🇪', // Flag emoji for Ireland
                            'FR': 'France 🇫🇷', // Flag emoji for France
                            'ES': 'Spain 🇪🇸', // Flag emoji for Spain
                            'CH': 'Switzerland 🇨🇭', // Flag emoji for Switzerland
                            'IL': 'Israel 🇮🇱', // Flag emoji for Israel
                            'AU': 'Australia 🇦🇺', // Flag emoji for Australia
                            'CA': 'Canada 🇨🇦', // Flag emoji for Canada
                            'BR': 'Brazil 🇧🇷' // Flag emoji for Brazil
                            // Add more mappings for other substrings as needed
                        };




                        ///// GET PART DATA ACROSS ROW//////




                        /// GET TYPE ////
                        let type = cell.parentElement.cells[16].textContent.trim();
                        // Add options for Motherboard, Memory, Cable, Dongle, Switch, ATS
                        if (type === "Power Supply") {
                            type = "Power Supply ⚡";
                        } else if (type === "Fiber Optic") {
                            type = "Fiber Optic ➰";
                        } else if (type === "Motherboard") {
                            type = "Motherboard 🖥️";
                        } else if (type === "Memory") {
                            type = "Memory 💾";
                        } else if (type === "Cable") {
                            type = "Cable 🔌";
                        } else if (type === "Dongle") {
                            type = "Dongle 🎟";
                        } else if (type === "Switch") {
                            type = "Switch 🖥";
                        } else if (type === "ATS") {
                            type = "ATS ⚡";
                        }




                        // Function to find column index by header name
                        function findColumnIndex(headerName) {
                            const headerRow = document.querySelector('table thead tr');
                            if (!headerRow) return -1;

                            const headers = headerRow.querySelectorAll('th');
                            for (let i = 0; i < headers.length; i++) {
                                if (headers[i].textContent.trim().toLowerCase() === headerName.toLowerCase()) {
                                    return i;
                                }
                            }
                            return -1; // Column not found
                        }

                        // Get the part ID and remove "📦" if present
                        let partIdIndex = findColumnIndex('Part Id');
                        let partId = partIdIndex !== -1 ? cell.parentElement.cells[partIdIndex].textContent.trim() : "null";
                        partId = partId === "null" ? "🛇" : partId.replace("📦", '');

                        // Create a button element with the ✏️ emoji as its text content
                        const partIdLink = partId === "null" ? "🛇" : `<button title="Edit Part" onclick="window.open('https://mobility.amazon.com/part/part/${partId}/edit?', '_blank')">✏️</button>`;

                        // Get the serial ID and remove " ( RMA )" if present
                        let serialIdIndex = findColumnIndex('Serial Id');
                        let serialId = serialIdIndex !== -1 ? cell.parentElement.cells[serialIdIndex].textContent.trim() : "null";
                        serialId = serialId === "null" ? "🛇" : serialId.replace(/\s*\(\s*RMA\s*\)$/, '');

                        // Create a hyperlink element
                        const serialIdLink = serialId === "null" ? "🛇" : `<a href="https://mobility.amazon.com/part/part/${partId}" target="_blank">${serialId}</a>`;

                        // GET MPN
                        let mpnIndex = findColumnIndex('MPN');
                        let mpn = mpnIndex !== -1 ? cell.parentElement.cells[mpnIndex].textContent.trim() : "null";
                        mpn = mpn === "null" ? "🛇" : mpn;

                        // GET RMA#
                        let rmaNumberIndex = findColumnIndex('RMA Number');
                        let rmaNumber = rmaNumberIndex !== -1 ? cell.parentElement.cells[rmaNumberIndex].textContent.trim() : "null";
                        rmaNumber = rmaNumber === "null" ? "🛇" : rmaNumber;

                        // GET BUILD ID
                        let buildIdIndex = findColumnIndex('Build Id');
                        let buildId = buildIdIndex !== -1 ? cell.parentElement.cells[buildIdIndex].textContent.trim() : "null";
                        buildId = buildId === "null" ? "🛇" : buildId;

                        // GET VENDOR LOGO
                        let vendorIndex = findColumnIndex('Vendor Name');
                        let vendor = vendorIndex !== -1 ? cell.parentElement.cells[vendorIndex].textContent.trim() : "UNKNOWN";
                        vendor = vendor === "UNKNOWN" ? "🛇" : vendor;


                        // Check if the vendor is Quanta
                        if (vendor === "Quanta") {
                            // Replace vendor with the image tag with smaller size
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/quanta-computer_Logo-removebg-preview.png" alt="Quanta" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Annapurna") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne%40/Vendor%20Logos/Annapurna_Labs.png" alt="Annapurna" style="max-width: 250px; max-height: 125px;">';
                        } else if (vendor === "Annapurna Labs") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne%40/Vendor%20Logos/Annapurna_Labs.png" alt="Annapurna" style="max-width: 250px; max-height: 125px;">';
                        } else if (vendor === "Accton") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Accton-Logo-Color.png" alt="Accton" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Vendor_Accton") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Accton-Logo-Color.png" alt="Accton" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Adva") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/ADVA_Logo.png" alt="Adva" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Juniper") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Juniper_Networks.png" alt="Juniper" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "JUNIPER NETWORKS") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Juniper_Networks.png" alt="Juniper" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "ZT") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/zt-systems-companyupdate-removebg-preview.png" alt="ZT" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "zt") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/zt-systems-companyupdate-removebg-preview.png" alt="ZT" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Foxconn") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Foxconn_Logo.svg.png" alt="Foxconn" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Vendor_Foxconn") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Foxconn_Logo.svg.png" alt="Vendor_Foxconn" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Synnex") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Synnex_Corporation_logo.svg.png" alt="Synnex" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "HYVE") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Hyve_Solutions.pngg" alt="Hyve" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Vendor_MITAC") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/MiTAC_logo.svg.png" alt="MITAC" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Amphenol") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Amphenol_Logo.svg.png" alt="Amphenol" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "AOI") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/AOI_LOGO.png" alt="AOI" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Asia Vital Components") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/AsiaVitalComponents.png" alt="Asia Vital Components" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "BIZLINK") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/BizLink_Logo_Logo-removebg-preview.png" alt="BIZLINK" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Celestica") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Celestica.png" alt="Celestica" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Citrix") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Citrix.png" style="max-width: 150px; max-height: 75px;">';

                        } else if (vendor === "Cisco") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Cisco_Logo.png" alt="Cisco" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "COOLJAG") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Cooljag.png" alt="COOLJAG" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Cavium") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Cavium_Logo.jpg" alt="Cavium" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Ciena") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Ciena_Logo.png" alt="Ciena" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Coriant") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Coriant_Logo.png" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "DELTA") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/DELTA_Electronics_Logo.png" alt="DELTA" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Delta") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/DELTA_Electronics_Logo.png" alt="DELTA" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "DELTA OPTICAL") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/DELTA_Electronics_Logo.png" alt="DELTA OPTICAL" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Flextronics") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Flextronics.png" alt="Flextronics" style="max-width: 200px; max-height: 110px;">';
                        } else if (vendor === "Vendor_Flextronics") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Flextronics.png" alt="Flextronics" style="max-width: 200px; max-height: 110px;">';
                        } else if (vendor === "Finisar") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/finisar-logo.png" alt="Finisar" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "FURUKAWA") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Furukawa.png" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "HPE") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/HPE.png" alt="HPE" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Hitachi") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Hitachi-Logo.png" alt="Hitachi" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Hynix") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Hynix.png" alt="Hynix" style="max-width: 150px; max-height: 75px;">';

                        } else if (vendor === "HYNIX") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Hynix.png" alt="HYNIX" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "HYVE") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Hyve_Solutions.png" alt="HYVE" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Infinera") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Infinera-Logo.png" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "INTEL") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Intel.png" alt="INTEL" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Intel") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Intel.png" alt="Intel" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "INNOLIGHT TECHNOLOGY") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Innolight.png" alt="INNOLIGHT TECHNOLOGY" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Jabil") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Jabil-Logo.wine.png" alt="Jabil" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "LiteON") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/liteon-logo-removebg-preview.png" alt="LiteON" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "LONGWELL") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Longwell_Logo.png" alt="LONGWELL" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Micron") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Micron-Loog.png" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "MITAC") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/MiTAC_logo.svg.png" alt="MITAC" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Molex") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Molex.png" alt="Molex" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "NetScout") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/NetScout_logo.png" alt="NetScout" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "NIDEC") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/nidec-logo-png-transparent.png" alt="NIDEC" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Nvidia") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Nvidia.png" alt="Nvidia" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "QUANTA AWS") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/quanta-computer_Logo-removebg-preview.png" alt="QUANTA AWS" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Raspberry Pi Foundation") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Raspberry_Pi.png" alt="Raspberry Pi Foundation" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Sanyo Denki") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Sanyo_logo.png" alt="Sanyo Denki" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Source Photonics") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Source-Photonics.png" alt="Source Photonics" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Sunon") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Sunon.png" alt="Sunon" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "TE Connectivity") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/TE_Connectivity_logo.svg.png" alt="TE Connectivity" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Western Digital") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Western_Digital_logo.svg.png" alt="Western Digital" style="max-width: 150px; max-height: 75px;">';

                        } else if (vendor === "Vendor_Synnex") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Synnex_Corporation_logo.svg.png" alt="Vendor_Synnex" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Vendor_Quanta") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/quanta-computer_Logo-removebg-preview.png" alt="Vendor_Quanta" style="max-width: 150px; max-height: 75px;">';

                        } else if (vendor === "Vendor_ZTSystems") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/ZT-Systems.png" alt="Vendor_ZTSystems" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "ZTSystems") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/ZT-Systems.png" alt="ZTSystems" style="max-width: 150px; max-height: 75px;">';


                        } else if (vendor === "Toshiba") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/TOSHIBA_Logo.png" alt="Toshiba" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Seagate") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Seagate_logo.png" alt="Seagate" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Samsung") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Samsung.png" alt="Seagate" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "SanDisk") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/SanDisk_Logo.png" alt="SanDisk" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "HP") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/HPE_Logo-removebg-preview.png" alt="HP" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Quantum") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Quantum_Corporation_logo.PNG" alt="Quantum" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "FUJIFILM") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Fujifilm-logo.png" alt="FUJIFILM" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Dell") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Dell_Logo.png" alt="Dell" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Kingston") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Kingston-Emblem.png" alt="Kingston" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Amazon") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Amazon_logo.png" alt="Amazon" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Panasonic") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Panasonic_logo_(Blue).PNG" alt="Panasonic" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Acbel") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Acbel.png" alt="Acbel" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Luxtera") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Luxtera-removebg-preview.png" alt="Luxtera" style="max-width: 150px; max-height: 75px;">';

                        } else if (vendor === "InVue") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/InVue_Logo.png" alt="InVue" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Vendor_Celestica") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Celestica.png" alt="Vendor_Celestica" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Oplink") {
                            vendor = '<img src="imageUrl" alt="Oplink" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Kaiam") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/KAIAM.png" alt="Kaiam" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "NetGear") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Netgear_logo.png" alt="NetGear" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Vendor_Kingston") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Kingston-Emblem.png" alt="Vendor_Kingston" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "WWT") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/World_Wide_Technology.png" alt="WWT" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Vendor_WWT") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/World_Wide_Technology.png" alt="WWT" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Vendor_Quantum") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Quantum_Corporation_logo.PNG" alt="Vendor_Quantum" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "ATP") {
                            vendor = '<img src="imageUrl" alt="ATP" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Oclaro") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/OCLARO.png" alt="Oclaro" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Vendor_Pegatron") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Pegatron.png" alt="Vendor_Pegatron" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "AMD") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/AMD-Logo.png" alt="AMD" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "NetIG") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/NetIG.png" alt="NetIG" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Vendor_Inventec") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Inventec.png" alt="Vendor_Inventec" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "FIBREFAB") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/fiberfab_logo.PNG" alt="FIBREFAB" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Lenovo") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Lenovo_Global_Corporate_Logo.png" alt="Lenovo" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "ColorChip") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Colorchip.png" alt="ColorChip" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Vendor_Panasonic") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Panasonic_logo_(Blue).PNG" alt="Vendor_Panasonic" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "APC") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/APC-Logo.PNG" alt="APC" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Vendor_Jabil") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Jabil-Logo.wine.png" alt="Vendor_Jabil" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Volex") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Volex-removebg-preview.png" alt="Volex" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Servertech") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/ServerTech.png" alt="Servertech" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Chenbro") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Chenbro_logo.png" alt="Chenbro" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "Vendor_Delta") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/DELTA_Electronics_Logo.png" alt="Vendor_Delta" style="max-width: 150px; max-height: 75px;">';
                        } else if (vendor === "SuperMicro") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Super_Micro_Computer_Logo.png" alt="SuperMicro" style="max-width: 150px; max-height: 75px;">';

                        } else if (vendor === "Network Integrity Systems") {
                            vendor = '<img src="https://drive-render.corp.amazon.com/view/pjbyrne@/Vendor%20Logos/Network-Integrity-Systems.png" alt="Network Integrity Systems" style="max-width: 150px; max-height: 75px;">';

                        } else {
                            // Default to the vendor text if there's no match
                        }






                        // GET STATE
                        let stateIndex = findColumnIndex('State');
                        let stateStatus = '';
                        const stateStatusText = stateIndex !== -1 ? cell.parentElement.cells[stateIndex].textContent.trim() : '';
                        switch (stateStatusText) {
                            case "RECEIVED":
                                stateStatus = "RECEIVED 📦";
                                break;
                            case "SPARE":
                                stateStatus = "SPARE 🔧";
                                break;
                            case "FOUND":
                                stateStatus = "FOUND 🔍";
                                break;
                            case "MISSING":
                                stateStatus = "MISSING ❌";
                                break;
                            case "PENDING_DESTRUCTION":
                                stateStatus = "PENDING_DESTRUCTION ⏳";
                                break;
                            case "IN_TRANSIT":
                                stateStatus = "IN_TRANSIT 🚚";
                                break;
                            case "BROKEN":
                                stateStatus = "BROKEN 💔";
                                break;
                            case "DEPLOYED":
                                stateStatus = "DEPLOYED 🚀";
                                break;
                            case "OUT_FOR_REPAIR":
                                stateStatus = "OUT_FOR_REPAIR 🔨";
                                break;
                            case "PENDING_SANITIZATION":
                                stateStatus = "PENDING SAN 🔬";
                                break;
                            case "PENDING_RMA":
                                stateStatus = "PENDING RMA ⌛";
                                break;
                            case "RMA_PROCESSED":
                                stateStatus = "RMA PROCESSED ✅";
                                break;
                            case "RMA":
                                stateStatus = "RMA 🚚";
                                break;
                            case "RESERVED_FOR_BUILD":
                                stateStatus = "RESERVED FOR BUILD 🛠️";
                                break;
                            case "DESTROYED":
                                stateStatus = "DESTROYED 💥";
                                break;
                            case "PENDING_WINSTON_WOLFE":
                                stateStatus = "PENDING_WINSTON_WOLFE 🕵️";
                                break;
                                // Add more cases for other state statuses as needed
                            default:
                                stateStatus = stateStatusText; // Use the status text itself for any other value
                        }



                        // GET CATEGORY
                        let categoryIndex = findColumnIndex('Category');
                        let categoryStatus = '';
                        const categoryStatusText = categoryIndex !== -1 ? cell.parentElement.cells[categoryIndex].textContent.trim() : '';
                        switch (categoryStatusText) {
                            case "CRITICAL_SPARE":
                                categoryStatus = "CRITICAL_SPARE ⚠️";
                                break;
                            case "DCO_SPARE":
                                categoryStatus = "DCO_SPARE 🛠️";
                                break;
                            case "BUILD":
                                categoryStatus = "BUILD 🧱";
                                break;
                            case "PROJECT":
                                categoryStatus = "PROJECT 🚧";
                                break;
                            default:
                                categoryStatus = categoryStatusText;
                        }

                        // GET ABSTRACT STATE
                        let abstractStateIndex = findColumnIndex('Abstract State');
                        let abstractState = '';
                        const abstractStateText = abstractStateIndex !== -1 ? cell.parentElement.cells[abstractStateIndex].textContent.trim() : '';
                        switch (abstractStateText) {
                            case "GONE":
                                abstractState = "GONE ❌";
                                break;
                            case "BIN":
                                abstractState = "BIN 🗑️";
                                break;
                            case "DEPLOYED":
                                abstractState = "DEPLOYED 🚀";
                                break;
                            default:
                                abstractState = abstractStateText;
                        }

                        // GET PO TEXT AND ADD HYPERLINK/FLAGS
                        let poIndex = findColumnIndex('PO Number');
                        let POText = poIndex !== -1 ? cell.parentElement.cells[poIndex].textContent.trim() : 'null';

                        let flag = '';

                        // Iterate through flag mappings and check for matches
                        for (const substring in flagMappings) {
                            if (POText.substring(0, substring.length) === substring) {
                                flag = flagMappings[substring];
                                break; // Exit loop once a match is found
                            }
                        }

                        let PO = POText === "null" ? "🛇" : `<a href="https://aws.argo.ocean-wave.aws.a2z.com/ordering/purchase_orders/${POText}" target="_blank">${POText}  ( ${flag} )</a>`;











                        ////📞🔧 FETCH PHONETOOL DATA /////


                        // Define the wait function
                        const wait = (ms) => {
                            return new Promise(resolve => setTimeout(resolve, ms));
                        };

                        // Define the xhr function to perform XMLHttpRequest
                        const xhr = (url, timeout, opt = {}) => {
                            Object.assign(opt, {
                                url,
                                timeout: timeout,
                                method: 'GET'
                            });
                            return new Promise((resolve, reject) => {
                                opt.onerror = opt.ontimeout = reject;
                                opt.onload = resolve;
                                GM_xmlhttpRequest(opt);
                            });
                        };

                        // Function to fetch user image from the specified URL
                        const fetchUserImage = async (username) => {
                            const url = `https://badgephotos.corp.AMAZON.com/?uid=${username}`;
                            const response = await xhr(url, 15000);
                            if (response && response.status === 200 && response.responseText) {
                                return response.finalUrl; // Return the final URL of the image
                            } else {
                                return null; // Return null if there's an error fetching the image
                            }
                        };

                        const fetchUserData = async (username) => {
                            const url = `https://phonetool.AMAZON.com/users/${username}`;
                            console.log('Fetching data from URL:', url);

                            return new Promise((resolve, reject) => {
                                GM_xmlhttpRequest({
                                    method: 'GET',
                                    url: url,
                                    onload: function(response) {
                                        if (response.status === 200 && response.responseText) {
                                            const parser = new DOMParser();
                                            const htmlDoc = parser.parseFromString(response.responseText, 'text/html');
                                            const contentElement = htmlDoc.getElementById('content');
                                            if (contentElement) {
                                                const employeeContainer = contentElement.querySelector('.employee-container');
                                                if (employeeContainer) {
                                                    const employeeData = employeeContainer.innerHTML.trim();
                                                    const fullNameMatch = employeeData.match(/targetUserName&quot;:&quot;(.*?)&quot;/);
                                                    const jobTitleMatch = employeeData.match(/targetUserJobTitle&quot;:&quot;(.*?)&quot;/);
                                                    const departmentNameMatch = employeeData.match(/targetUserDepartmentName&quot;:&quot;(.*?)&quot;/);
                                                    const buildingMatch = employeeData.match(/targetUserBuilding&quot;:&quot;(.*?)&quot;/);
                                                    const mobileNumberMatch = employeeData.match(/targetUserMobileNumber&quot;:&quot;(.*?)&quot;/);

                                                    const userData = {
                                                        fullName: fullNameMatch && fullNameMatch[1] ? fullNameMatch[1] : 'Missing',
                                                        jobTitle: jobTitleMatch && jobTitleMatch[1] ? jobTitleMatch[1] : 'Missing',
                                                        departmentName: departmentNameMatch && departmentNameMatch[1] ? departmentNameMatch[1] : 'Missing',
                                                        building: buildingMatch && buildingMatch[1] ? buildingMatch[1] : 'Missing',
                                                        mobileNumber: mobileNumberMatch && mobileNumberMatch[1] ? mobileNumberMatch[1] : 'Missing'
                                                    };

                                                    resolve(userData);
                                                } else {
                                                    reject('Employee container element not found in response');
                                                }
                                            } else {
                                                reject('Content element not found in response');
                                            }
                                        } else {
                                            reject('Error fetching user data');
                                        }
                                    },
                                    onerror: function(error) {
                                        reject(error);
                                    }
                                });
                            });
                        };



                        /////📞 DISPLAY PHONETOOL CALLING CARD /////

                        const displayCallingCard = async (username) => {
                            try {
                                const imageSrc = await fetchUserImage(username);
                                if (!imageSrc) {
                                    throw new Error('Error fetching user image');
                                }
                                console.log('Image fetched from URL:', imageSrc);

                                const userData = await fetchUserData(username);
                                if (!userData) {
                                    throw new Error('Error fetching user data');
                                }
                                console.log('User Data:', userData);

                                // Remove any existing calling card
                                const existingCallingCard = document.querySelector('.calling-card');
                                if (existingCallingCard) {
                                    existingCallingCard.remove();
                                }

                                // Create the calling card container
                                const callingCard = document.createElement('div');
                                callingCard.id = 'calling-card'; // Set id for easy removal
                                callingCard.classList.add('calling-card'); // Add a class for styling
                                callingCard.style.position = 'fixed';
                                callingCard.style.top = '50%';
                                callingCard.style.left = '50%';
                                callingCard.style.transform = 'translate(-50%, -50%)';
                                callingCard.style.zIndex = '9999';
                                callingCard.style.backgroundColor = '#85c1e9'; // Background color similar to the sample
                                callingCard.style.color = '#333'; // Text color
                                callingCard.style.padding = '20px';
                                callingCard.style.cursor = 'default';
                                callingCard.style.borderRadius = '20px';
                                callingCard.style.boxShadow = '0 0 20px rgba(0, 0, 0, 0.5)'; // Box shadow similar to the sample
                                callingCard.style.border = '4px solid #ccc'; // Lighter border color
                                callingCard.style.borderTop = '4px solid #f5f5f5'; // Lighter color on top for shine effect
                                callingCard.style.borderLeft = '4px solid #f5f5f5'; // Lighter color on left for shine effect
                                callingCard.style.backgroundClip = 'padding-box'; // Apply background within padding
                                callingCard.style.fontFamily = 'Roboto, Arial, sans-serif'; // Set font family with fallbacks
                                callingCard.style.fontWeight = 'bold'; // Bold text
                                callingCard.style.textShadow = '1px 1px 2px rgba(255, 255, 255, 0.5)'; // Soft text shadow for 3D effect
                                callingCard.style.display = 'flex'; // Flexbox for layout
                                callingCard.style.alignItems = 'center'; // Center items vertically

                                // Create and append the close button to the calling card
                                const closeButton = document.createElement('button');
                                closeButton.textContent = '❌'; // Set the button text content to the delete emoji directly
                                closeButton.style.position = 'absolute';
                                closeButton.style.top = '10px';
                                closeButton.style.right = '10px';
                                closeButton.style.color = 'white'; // Text color
                                closeButton.style.border = 'none';
                                closeButton.style.background = 'none'; // Remove background color
                                closeButton.style.width = '20px';
                                closeButton.style.height = '20px';
                                closeButton.style.cursor = 'pointer';
                                callingCard.appendChild(closeButton);

                                closeButton.addEventListener('click', () => {
                                    callingCard.remove();
                                });


                                // Create and append the image element
                                const imageElement = document.createElement('img');
                                imageElement.setAttribute('src', imageSrc);
                                imageElement.setAttribute('id', 'badge-photo');
                                imageElement.style.borderRadius = '5px';
                                imageElement.style.maxWidth = '100px'; // Adjust the size as needed
                                imageElement.style.maxHeight = '100px'; // Adjust the size as needed
                                imageElement.style.marginRight = '20px'; // Spacing between image and user data
                                imageElement.style.border = '1px solid black'; // Add black border
                                callingCard.appendChild(imageElement);

                                // Create and append the user data elements using inline styles for styled text
                                const userDataElement = document.createElement('div');
                                userDataElement.innerHTML = `
           <div style="margin-bottom: 5px;">
                 <span style="font-weight: bold; font-size: 14px;"></span>
                 <span style="font-size: 14px; background-color: white; border-radius: 5px; padding: 1.25px 2.5px; font-family: 'Arial', sans-serif; border: 1px solid black;">
                        ${userData.fullName} (${username})
                 </span>
           </div>

            <div>
                <span style="font-weight: bold; font-size: 14px;">👷 </span> <span style="font-size: 14px; font-family: 'Arial'">${userData.jobTitle}</span>
            </div>
            <div>
                <span style="font-weight: bold; font-size: 14px;">💼 </span> <span style="font-size: 14px; font-family: 'Arial'">${userData.departmentName}</span>
            </div>
            <div>
                <span style="font-weight: bold; font-size: 14px;">🏬 </span> <span style="font-size: 14px; font-family: 'Arial'">${userData.building}</span>
            </div>
            <div>
                <span style="font-weight: bold; font-size: 14px;">📱 </span> <span style="font-size: 14px; font-family: 'Arial'">${userData.mobileNumber ? userData.mobileNumber : 'Missing'}</span>
            </div>
            <div>
                 <span style="font-weight: bold; font-size: 14px;">📧 </span> <span style="font-size: 14px; font-family: 'Arial'" id="emailAddress">${username}@amazon.com</span>
            </div>
        `;
                                callingCard.appendChild(userDataElement);

                                // Append the calling card to the document body
                                document.body.appendChild(callingCard);
                            } catch (error) {
                                console.error(error.message);
                            }
                        };

                        // Initialize the calling card functionality
                        const init = () => {
                            // Select the usernames and create hyperlinks
                            const usernames = document.querySelectorAll('table.table-bordered.table-striped.table-condensed tbody tr td:nth-child(22)');

                            usernames.forEach(username => {
                                const link = document.createElement('a');
                                link.classList.add('username-link'); // Add class for styling if needed
                                link.href = '#'; // Set href to '#' for placeholder
                                link.textContent = username.textContent.trim(); // Set the link text to the username

                                link.addEventListener('click', async (event) => {
                                    event.preventDefault(); // Prevent default link behavior

                                    const usernameText = link.textContent.trim();
                                    if (usernameText !== 'null' && usernameText !== '') {
                                        await displayCallingCard(usernameText);
                                    }
                                });

                                // Replace the username text with the hyperlink
                                username.innerHTML = '';
                                username.appendChild(link);
                            });
                        };

                        // Call the init function to initialize the calling card functionality
                        init();







                        ///// 📦 GET RECEIVED BY DATA //////

                        let usernameLink = '';

                        // Fetch data and create the username link
                        const fetchAndDisplayUsername = async (partId) => {
                            try {
                                const url = `https://mobility.AMAZON.com/part/part/${partId}`;
                                console.log('Accessing URL:', url);

                                const response = await fetch(url);
                                const htmlText = await response.text();

                                console.log('ReceivedBy HTML response received:', htmlText);

                                // Create a temporary div to parse the HTML response
                                const tempDiv = document.createElement('div');
                                tempDiv.innerHTML = htmlText;

                                console.log('Parsed HTML:', tempDiv);

                                // Find the td element containing the username
                                const usernameElement = tempDiv.querySelector('td:nth-child(4)');
                                if (!usernameElement) {
                                    console.log('Username element not found.');
                                    return; // Exit the function if username element is not found
                                }

                                // Extract the username from the td element
                                const usernameText = usernameElement.textContent.trim();
                                const usernameMatch = usernameText.match(/\(([^)]+)\)/);
                                const username = usernameMatch ? usernameMatch[1] : '';

                                console.log('Extracted username:', username);

                                // Create the username link
                                usernameLink = document.createElement('a');
                                usernameLink.textContent = username;
                                usernameLink.href = '#';
                                usernameLink.title = '📞🃏 Click to open Calling Card'; // Add hover title


                                console.log('Created username link:', usernameLink);

                                // Add click event listener to the hyperlink
                                usernameLink.addEventListener('click', () => {
                                    displayCallingCard(username); // Call the displayCallingCard function with the username
                                });

                                // Append the username link to the DOM
                                const receivedBySpan = document.getElementById('receivedBySpan');
                                if (receivedBySpan) {
                                    receivedBySpan.appendChild(usernameLink);
                                } else {
                                    console.error('Element with ID "receivedBySpan" not found.');
                                }

                            } catch (error) {
                                console.error('Error fetching or processing data:', error);
                            }
                        };

                        // Call the fetchAndDisplayUsername function with the extracted partId
                        fetchAndDisplayUsername(partId);






                        // GET USER CUSTODY DATA
                        let userCustodyIndex = findColumnIndex('User Custody');
                        let userCustody = '';

                        if (userCustodyIndex !== -1) {
                            let userCustodyText = cell.parentElement.cells[userCustodyIndex].textContent.trim();
                            userCustody = userCustodyText === "nullull"
                                ? "🛇"
                            : `<a href="#" class="userCustodyLink">${userCustodyText}</a>`;
                        } else {
                            userCustody = "🛇"; // Column not found
                        }

                        // Add event listener to userCustodyLink class
                        document.addEventListener('click', async (event) => {
                            if (event.target.classList.contains('userCustodyLink')) {
                                event.preventDefault(); // Prevent default link behavior

                                const username = event.target.textContent.trim();
                                if (username !== 'null' && username !== '') {
                                    await displayCallingCard(username);
                                }
                            }
                        });






                        ////🚀   GET DEPLOYED BY DATA //////

                        // Define the deployedUserLink variable
                        let deployedUserLink = '';

                        // Fetch data and create the deployed user link
                        const fetchDeployedUser = async (partId) => {
                            try {
                                const url = `https://mobility.AMAZON.com/part/part/${partId}`;
                                console.log('Accessing URL:', url);

                                const response = await fetch(url, {
                                    method: 'GET',
                                    headers: {
                                        'Content-Type': 'text/html',
                                    },
                                });

                                const htmlText = await response.text();
                                console.log('DeployedBy HTML response received:', htmlText);

                                // Create a temporary div to parse the HTML response
                                const tempDiv = document.createElement('div');
                                tempDiv.innerHTML = htmlText;

                                console.log('Parsed HTML:', tempDiv);

                                // Find the sort table
                                const sortTable = tempDiv.querySelector('#sortTable');

                                if (!sortTable) {
                                    console.log('Table with ID "sortTable" not found.');
                                    return;
                                }

                                console.log('Table with ID "sortTable" found:', sortTable);

                                // Find the correct row containing "DEPLOYED"
                                const rows = sortTable.querySelectorAll('tbody tr');
                                let deployedUsername = '';
                                let deployedDate = ''; // Variable to store the date

                                rows.forEach((row) => {
                                    const cells = row.querySelectorAll('td');
                                    if (cells[2].textContent.trim() === 'DEPLOYED') {
                                        // Extract the username
                                        const userInfo = cells[3].textContent.trim();
                                        const matches = userInfo.match(/\(([^)]+)\)/);
                                        if (matches && matches.length > 1) {
                                            deployedUsername = matches[1];
                                        }

                                        // Extract the date two cells before
                                        deployedDate = cells[0].textContent.trim();
                                    }
                                });

                                if (!deployedUsername) {
                                    console.log('Text "DEPLOYED" not found in the table or username format incorrect.');
                                    return;
                                }

                                if (!deployedDate) {
                                    console.log('Deployed date not found.');
                                    return;
                                }

                                console.log('Extracted deployed username:', deployedUsername);
                                console.log('Extracted deployed date:', deployedDate);

                                // Calculate the difference in days
                                const today = new Date();
                                const deployedDateObject = new Date(deployedDate);
                                const daysAgo = Math.ceil((today - deployedDateObject) / (1000 * 60 * 60 * 24));

                                console.log('Days ago:', daysAgo);

                                // Format the date (optional, if needed)
                                const formattedDate = deployedDateObject.toLocaleDateString('en-US', {
                                    year: 'numeric',
                                    month: 'short',
                                    day: 'numeric',
                                });

                                // Create the deployed user link
                                deployedUserLink = document.createElement('a');
                                deployedUserLink.textContent = deployedUsername;
                                deployedUserLink.href = '#';
                                deployedUserLink.title = '📞🃏 Click to open Calling Card'; // Add hover title

                                console.log('Created deployed user link:', deployedUserLink);

                                // Add click event listener to the deployedUserLink
                                deployedUserLink.addEventListener('click', () => {
                                    displayCallingCard(deployedUsername); // Call the displayCallingCard function with the deployed username
                                });

                                // Update the deployedBySpan element with the deployed user link, date, and days ago
                                const deployedBySpan = document.getElementById('deployedBySpan');
                                if (deployedBySpan) {
                                    deployedBySpan.textContent = `${formattedDate} (${daysAgo} days ago) by `;
                                    deployedBySpan.appendChild(deployedUserLink);
                                } else {
                                    console.log('Element with ID "deployedBySpan" not found.');
                                }
                            } catch (error) {
                                console.error('Error fetching or processing data:', error);
                            }
                        };

                        // Call the fetchDeployedUser function with the extracted partId
                        fetchDeployedUser(partId);








                        ////🗓️ LAST UPDATED USER //////

                        // Define the lastUpdatedUserLink variable
                        let lastUpdatedUserLink = '';

                        // Fetch data and create the last updated user link
                        const fetchLastUpdatedUser = async (partId) => {
                            try {
                                const url = `https://mobility.AMAZON.com/part/part/${partId}`;
                                console.log('Accessing URL:', url);

                                const response = await fetch(url, {
                                    method: 'GET',
                                    headers: {
                                        'Content-Type': 'text/html',
                                    },
                                });

                                const htmlText = await response.text();
                                console.log('LastUser HTML response received:', htmlText);

                                // Create a temporary div to parse the HTML response
                                const tempDiv = document.createElement('div');
                                tempDiv.innerHTML = htmlText;

                                console.log('Parsed HTML:', tempDiv);

                                // Find the sort table
                                const sortTable = tempDiv.querySelector('#sortTable');

                                if (!sortTable) {
                                    console.log('Table with ID "sortTable" not found.');
                                    return;
                                }

                                console.log('Table with ID "sortTable" found:', sortTable);

                                // Get the last row of the table
                                const rows = sortTable.querySelectorAll('tbody tr');
                                const lastRow = rows[rows.length - 1];

                                if (!lastRow) {
                                    console.log('Last row not found in the table.');
                                    return;
                                }

                                // Get the username from the last row
                                const cells = lastRow.querySelectorAll('td');
                                const userInfo = cells[3].textContent.trim();
                                const matches = userInfo.match(/\(([^)]+)\)/);

                                if (matches && matches.length > 1) {
                                    const lastUpdatedUsername = matches[1];
                                    console.log('Last updated username:', lastUpdatedUsername);

                                    // Create the last updated user link
                                    lastUpdatedUserLink = document.createElement('a');
                                    lastUpdatedUserLink.textContent = lastUpdatedUsername;
                                    lastUpdatedUserLink.href = '#';
                                    lastUpdatedUserLink.title = '📞🃏 Click to open Calling Card'; // Add hover title



                                    console.log('Created last updated user link:', lastUpdatedUserLink);

                                    // Add click event listener to the lastUpdatedUserLink
                                    lastUpdatedUserLink.addEventListener('click', () => {
                                        displayCallingCard(lastUpdatedUsername); // Call the displayCallingCard function with the last updated username
                                    });

                                    // Append the last updated user link after the last updated date
                                    const lastUpdatedSpan = document.getElementById('lastUpdatedSpan');
                                    if (lastUpdatedSpan) {
                                        lastUpdatedSpan.appendChild(document.createTextNode(' by '));
                                        lastUpdatedSpan.appendChild(lastUpdatedUserLink);
                                    } else {
                                        console.log('Element with ID "lastUpdatedSpan" not found.');
                                    }
                                } else {
                                    console.log('Username format incorrect in the last row.');
                                }
                            } catch (error) {
                                console.error('Error fetching or processing data:', error);
                            }
                        };

                        // Call the fetchLastUpdatedUser function with the extracted partId
                        fetchLastUpdatedUser(partId);











                        ///🗓️ GET RECEIVED / UPDATED DATES /////



                        // Get the indices for 'Last Updated' and 'Received Date' columns
                        const lastUpdatedIndex = findColumnIndex('Last Updated Time');
                        const receivedIndex = findColumnIndex('Received At');

                        // Get and convert the dates
                        const lastUpdated = lastUpdatedIndex !== -1 ? convertDate(cell.parentElement.cells[lastUpdatedIndex].textContent.trim()) : '';
                        const received = receivedIndex !== -1 ? convertDate(cell.parentElement.cells[receivedIndex].textContent.trim()) : '';

                        // Convert lastUpdated and received dates to days ago
                        const today = new Date();
                        const lastUpdatedDate = isValidDate(lastUpdated) ? new Date(lastUpdated) : null;
                        const receivedDate = isValidDate(received) ? new Date(received) : null;
                        const lastUpdatedDays = lastUpdatedDate ? Math.ceil((today - lastUpdatedDate) / (1000 * 60 * 60 * 24)) : '';
                        const receivedDays = receivedDate ? Math.ceil((today - receivedDate) / (1000 * 60 * 60 * 24)) : '';




                        // Create and style the popup form
                        const popupForm = document.createElement('div');
                        popupForm.className = 'popup-form'; // Add a class for easier targeting
                        popupForm.style.backgroundColor = '#d4e6f1'; // Semi-transparent white background
                        popupForm.style.color = '#333'; // Dark text color
                        popupForm.style.padding = '20px';
                        popupForm.style.cursor = 'default';
                        popupForm.style.borderRadius = '20px'; // More rounded corners
                        popupForm.style.boxShadow = '0 0 20px rgba(0, 0, 0, 0.5)'; // Thicker and more pronounced shadow
                        popupForm.style.position = 'fixed';
                        popupForm.style.zIndex = '9999';
                        popupForm.style.top = '50%';
                        popupForm.style.left = 'calc(80% - 100px)'; // Adjust the distance from the right side here
                        popupForm.style.transform = 'translate(-50%, -50%)';
                        popupForm.style.width = '450px'; // Adjust width here
                        popupForm.style.transition = 'opacity 0.3s ease';
                        popupForm.style.border = '4px solid #ccc'; // Lighter border color
                        popupForm.style.borderTop = '4px solid #f5f5f5'; // Lighter color on top for shine effect
                        popupForm.style.borderLeft = '4px solid #f5f5f5'; // Lighter color on left for shine effect
                        popupForm.style.backgroundClip = 'padding-box'; // Apply background within padding
                        popupForm.style.fontFamily = 'Roboto, Arial, sans-serif'; // Set font family with fallbacks
                        popupForm.style.fontWeight = 'bold'; // Bold text
                        popupForm.style.textShadow = '1px 1px 2px rgba(255, 255, 255, 0.5)'; // Soft text shadow for 3D effect

                        document.body.appendChild(popupForm); // Append the popup form to the document body





                        // Create close button (red X emoji)
                        const closeButton = document.createElement('span');
                        closeButton.innerHTML = '❌'; // Red X emoji
                        closeButton.style.position = 'absolute';
                        closeButton.style.top = '10px';
                        closeButton.style.right = '10px';
                        closeButton.style.fontSize = '20px';
                        closeButton.style.color = 'red';
                        closeButton.style.cursor = 'pointer';
                        closeButton.addEventListener('click', function() {
                            popupForm.remove();
                        });










                        ////POPUP FORM CONTENT /////
                        popupForm.innerHTML = `

    <div style="margin-bottom: 10px; margin-top: -10px; padding-bottom: 5px; text-align: center;">
    <span style="font-weight: bold; font-size: 20px; color: navy; padding: 5px 10px; border-radius: 5px;"></span><span style="font-size: 16px; color: black; padding: 5px 10px; border-radius: 5px; font-family: 'Courier New', Courier, monospace;">${vendor}</span>
</div>

<div style="margin-bottom: 10px; padding-bottom: 5px;">
    <span style="font-weight: bold; font-family: 'Roboto', Arial, sans-serif; font-size: 16px; color: navy; padding: 5px 10px; border-radius: 5px;">▌│█║▌║▌║</span><span style="font-size: 14px; background-color: white; color: black; padding: 5px 10px; border-radius: 5px; font-family: 'Arial', sans-serif; border: 1px solid black;">${serialIdLink} ${partIdLink}</span>
</div>

<div style="margin-bottom: 10px; padding-bottom: 5px;">
    <span style="font-weight: bold; font-size: 16px; color: navy; padding: 5px 10px; border-radius: 5px;">🔡 TYPE </span><span style="font-size: 14px; background-color: white; color: black; padding: 5px 10px; border-radius: 5px; font-family: 'Arial', sans-serif; border: 1px solid black;">${type}</span>
</div>

<div style="margin-bottom: 10px; padding-bottom: 5px;">
    <span style="font-weight: bold; font-size: 16px; color: navy; padding: 5px 10px; border-radius: 5px;">🆔 MODEL </span><span style="font-size: 14px; background-color: white; color: black; padding: 5px 10px; border-radius: 5px; font-family: 'Arial', sans-serif; border: 1px solid black;">${mpn}</span>
</div>

<div style="margin-bottom: 10px; padding-bottom: 5px;">
    <span style="font-weight: bold; font-size: 16px; color: navy; padding: 5px 10px; border-radius: 5px;">📑 STATES </span><span style="font-size: 14px; background-color: white; color: black; padding: 5px 10px; border-radius: 5px; font-family: 'Arial', sans-serif; border: 1px solid black;">${stateStatus} / ${abstractStateText}</span>
</div>

<div style="margin-bottom: 10px; padding-bottom: 5px;">
    <span style="font-weight: bold; font-size: 16px; color: navy; padding: 5px 10px; border-radius: 5px;">🗂️ CATEGORY </span><span style="font-size: 14px; background-color: white; color: black; padding: 5px 10px; border-radius: 5px; font-family: 'Arial', sans-serif; border: 1px solid black;">${categoryStatus}</span>
</div>

<div style="margin-bottom: 10px; padding-bottom: 5px;">
    <span style="font-weight: bold; font-size: 16px; color: navy; padding: 5px 10px; border-radius: 5px;">📝 PO # </span><span style="font-size: 14px; background-color: white; color: black; padding: 5px 10px; border-radius: 5px; font-family: 'Arial', sans-serif; border: 1px solid black;">${PO}</span>
</div>

<div style="margin-bottom: 10px; padding-bottom: 5px;">
    <span style="font-weight: bold; font-size: 16px; color: navy; padding: 5px 10px; border-radius: 5px;">↩️ RMA # </span><span style="font-size: 14px; background-color: white; color: black; padding: 5px 10px; border-radius: 5px; font-family: 'Arial', sans-serif; border: 1px solid black;">${rmaNumber}</span>
</div>

<div style="margin-bottom: 10px; padding-bottom: 5px;">
    <span style="font-weight: bold; font-size: 16px; color: navy; padding: 5px 10px; border-radius: 5px;">🧱 BUILD ID </span><span style="font-size: 14px; background-color: white; color: black; padding: 5px 10px; border-radius: 5px; font-family: 'Arial', sans-serif; border: 1px solid black;">${buildId}</span>
</div>

<div style="margin-bottom: 10px; padding-bottom: 5px;">
    <span style="font-weight: bold; font-size: 16px; color: navy; padding: 5px 10px; border-radius: 5px;">📥 INJECTED </span>
    <span id="receivedBySpan" style="font-size: 14px; background-color: white; color: black; padding: 5px 10px; border-radius: 5px; font-family: 'Arial', sans-serif; border: 1px solid black;">${received} (${receivedDays ? `${receivedDays} days ago)by ` : ''}</span>
</div>
<div style="margin-bottom: 10px; padding-bottom: 5px;">
    <span style="font-weight: bold; font-size: 16px; color: navy; padding: 5px 10px; border-radius: 5px;">🚀 DEPLOYED </span>
    <span id="deployedBySpan" style="font-size: 14px; background-color: white; color: black; padding: 5px 10px; border-radius: 5px; font-family: 'Arial', sans-serif; border: 1px solid black;"></span>
</div>


<div style="margin-bottom: 10px; padding-bottom: 5px;">
    <span style="font-weight: bold; font-size: 16px; color: navy; padding: 5px 10px; border-radius: 5px;">📦 CUSTODY </span><span style="font-size: 14px; background-color: white; color: black; padding: 5px 10px; border-radius: 5px; font-family: 'Arial', sans-serif; border: 1px solid black;">${userCustody}</span>
</div>

<div style="margin-bottom: 10px; padding-bottom: 5px;">
    <span style="font-weight: bold; font-size: 16px; color: navy; padding: 5px 10px; border-radius: 5px;">🗓️ UPDATED: </span>
    <span id="lastUpdatedSpan" style="font-size: 14px; background-color: white; color: black; padding: 5px 10px; border-radius: 5px; font-family: 'Arial', sans-serif; border: 1px solid black;">${lastUpdated} (${lastUpdatedDays} days ago)</span>
</div>



`;


                        // Append the popup form to the document body
                        document.body.appendChild(popupForm);
                        // Append the close button to the popup form
                        popupForm.appendChild(closeButton);


                        // Update the active cell
                        activeCell = cell;

                    }
                });


            });
        }

        // Call the function to handle hover and click on cells containing serial numbers
        handleSerialNumberCellEvents();

        // Function to convert date string to a user-friendly format
        function convertDate(dateString) {
            const dateParts = dateString.split(' ');
            const date = dateParts[0]; // Get the date part
            const parsedDate = new Date(date);
            const options = { year: '2-digit', month: 'short', day: '2-digit' };
            const formattedDate = parsedDate.toLocaleDateString('en-US', options);
            return formattedDate.toUpperCase(); // Convert to uppercase
        }

        // Function to check if a date string is valid
        function isValidDate(dateString) {
            const date = new Date(dateString);
            return !isNaN(date.getTime());
        }
    })();




    const bulkEditButton = document.getElementById('part_search_bulk_edit_popup_button');

    if (bulkEditButton) {
        bulkEditButton.textContent = 'Bulk Edit'; // Replace the text content of the button
    } else {
        console.error('Bulk edit button not found.');
    }












    /////⬆️ SCROLL TOP PAGE BUTTON //////


    // Create the scroll-to-top element
    const scrollToTopButton = document.createElement('div');
    scrollToTopButton.innerHTML = '⬆️';
    scrollToTopButton.style.display = 'none';
    scrollToTopButton.style.position = 'fixed';
    scrollToTopButton.style.bottom = '20px';
    scrollToTopButton.style.left = '20px';
    scrollToTopButton.style.fontSize = '30px';
    scrollToTopButton.style.cursor = 'pointer';
    scrollToTopButton.style.zIndex = '1000';

    // Create the scroll toggle element
    const scrollToggleButton = document.createElement('div');
    scrollToggleButton.innerHTML = '↔️';
    scrollToggleButton.style.display = 'none';
    scrollToggleButton.style.position = 'fixed';
    scrollToggleButton.style.bottom = '20px';
    scrollToggleButton.style.left = '60px'; // Adjusted left position to appear after scroll up icon
    scrollToggleButton.style.fontSize = '30px';
    scrollToggleButton.style.cursor = 'pointer';
    scrollToggleButton.style.zIndex = '1000';

    // Append the scroll-to-top button to the body
    document.body.appendChild(scrollToTopButton);

    // Append the scroll toggle button to the body right after the scroll-to-top button
    document.body.insertBefore(scrollToggleButton, scrollToTopButton.nextSibling);


    // Variable to store the timeout
    let scrollTimeout;

    // Function to handle hiding the button after scrolling stops
    function hideScrollButton() {
        scrollToTopButton.style.display = 'none';
        scrollToggleButton.style.display = 'none';
    }

    // Show the scroll-to-top button and scroll toggle button when scrolling starts
    window.addEventListener('scroll', function() {
        scrollToTopButton.style.display = 'block';
        scrollToggleButton.style.display = 'block';

        // If there's a previous timeout, clear it
        if (scrollTimeout) {
            clearTimeout(scrollTimeout);
        }

        // Set a new timeout to hide the buttons after 5 seconds
        scrollTimeout = setTimeout(hideScrollButton, 5000);
    });

    // Smooth scroll to top when the button is clicked
    scrollToTopButton.addEventListener('click', function() {
        window.scrollTo({
            top: 0,
            behavior: 'smooth'
        });
    });

    // Scroll toggle functionality
    let isScrollingRight = true; // Initial state

    scrollToggleButton.addEventListener('click', function() {
        if (isScrollingRight) {
            window.scrollTo({
                left: document.body.scrollWidth,
                behavior: 'smooth'
            });
            isScrollingRight = false;
        } else {
            window.scrollTo({
                left: 0,
                behavior: 'smooth'
            });
            isScrollingRight = true;
        }
    });








    //////⬅️⬆️⬇️ SLIDE ELEMENTS /////////////////


    ///// ↖️ COMBINED SLIDE-IN/OUT BUTTON //////
    function toggleHideElements() {
        console.log('🔍 Initializing toggleHideElements function.');

        // Select the elements to slide
        const refinementsElement = document.querySelector('.search_refinements');
        const rowLeftElement = document.querySelector('.row-left7'); // Get the row-left7 element
        const targetElement = document.querySelector('.marginLeft10'); // Get the search bar element

        // Log to check if elements are found
        if (!refinementsElement) {
            console.log('❌ refinementsElement (.search_refinements) not found.');
        } else {
            console.log('🎯 refinementsElement found:', refinementsElement);
        }

        if (!rowLeftElement) {
            console.log('❌ rowLeftElement (.row-left7) not found.');
        } else {
            console.log('🎯 rowLeftElement found:', rowLeftElement);
        }

        if (!targetElement) {
            console.log('❌ targetElement (.marginLeft10) not found.');
        } else {
            console.log('🎯 targetElement found:', targetElement);
        }

        // Create a span element for the combined emoji
        const hideElementsEmoji = document.createElement('span');
        let isHidden = true; // Keep track of the button state
        const arrowIcon = '↖️'; // Combined arrow icon for the button
        const downIcon = '↘️'; // Downward diagonal arrow for the expanded state
        hideElementsEmoji.innerHTML = arrowIcon; // Initial state with combined arrow icon
        hideElementsEmoji.className = 'emoji'; // Add a class for styling
        hideElementsEmoji.style.marginLeft = '-10px';
        hideElementsEmoji.id = 'toggle_hide_elements_button'; // Assign an ID for targeting
        hideElementsEmoji.style.fontSize = '18px'; // Increase font size
        hideElementsEmoji.style.cursor = 'pointer'; // Set cursor to pointer
        hideElementsEmoji.title = 'Toggle Search Refinements and Search Bar'; // Add tooltip message

        // Add click event listener to the combined emoji
        hideElementsEmoji.addEventListener('click', function() {
            if (refinementsElement && rowLeftElement && targetElement) {
                // Toggle both elements' visibility
                refinementsElement.classList.toggle('hidden');
                rowLeftElement.classList.toggle('slide'); // Toggle the slide class on the row-left7 element

                const isVisible = targetElement.style.display !== 'none';
                if (isVisible) {
                    targetElement.style.transition = 'height 0.5s, padding 0.5s';
                    targetElement.style.height = '0';
                    targetElement.style.paddingTop = '0';
                    targetElement.style.paddingBottom = '0';
                    targetElement.style.overflow = 'hidden';
                    setTimeout(() => {
                        targetElement.style.display = 'none';
                    }, 500);
                } else {
                    targetElement.style.display = 'block';
                    targetElement.style.height = 'auto';
                    const height = targetElement.clientHeight + 'px';
                    targetElement.style.height = '0';
                    targetElement.style.transition = 'height 0.5s, padding 0.5s';
                    targetElement.style.paddingTop = '';
                    targetElement.style.paddingBottom = '';
                    setTimeout(() => {
                        targetElement.style.height = height;
                    }, 0);
                }

                // Toggle the emoji and title based on the state
                isHidden = !isHidden; // Toggle the button state
                hideElementsEmoji.innerHTML = isHidden ? arrowIcon : downIcon; // Change icon based on state
                hideElementsEmoji.title = isHidden ? 'Toggle Search Refinements and Search Bar' : 'Toggle Search Refinements and Search Bar';
            } else {
                console.log('⚠️ Cannot toggle, elements not found.');
            }
        });

        // Add hover effect to grow the emoji
        hideElementsEmoji.addEventListener('mouseover', function() {
            hideElementsEmoji.style.transform = 'scale(1.1)'; // Grow emoji on hover
        });

        // Remove hover effect to revert to original size
        hideElementsEmoji.addEventListener('mouseout', function() {
            hideElementsEmoji.style.transform = 'scale(1)'; // Revert to original size
        });

        // Get the parent div
        const parentDiv = document.querySelector('.margin5');

        // Append the toggle emoji before the Bulk Edit Parts button
        const bulkEditButton = document.getElementById('part_search_bulk_edit_popup_button');
        parentDiv.insertBefore(hideElementsEmoji, bulkEditButton);

        // Set transition timing for slide effect on refinementsElement
        const transitionDuration = '0.5s'; // Adjust as needed
        refinementsElement.style.transition = `height ${transitionDuration}, padding ${transitionDuration}`;

        console.log('✅ hideElementsEmoji created and appended.');
    }

    // Initialize the combined slide toggle function
    toggleHideElements();

    ///// 🔒 MORE TOOLS TAB //////
    let tabState = false;

    function toggleFormContainer() {
        const formContainer = document.getElementById('formContainer');
        formContainer.classList.toggle('active');
        moreToolsTab.textContent = tabState ? '🔒' : '🔓';
        tabState = !tabState;
    }

    // Create the tab element
    const moreToolsTab = document.createElement('div');
    moreToolsTab.id = 'customTab';
    moreToolsTab.textContent = '🔒';
    moreToolsTab.addEventListener('click', toggleFormContainer);
    moreToolsTab.setAttribute('title', '⚠️ Click for more tools! 🛠️'); // Tooltip added here
    moreToolsTab.style.marginLeft = '0px'; // Add left margin to pull it closer to the title text
    moreToolsTab.style.marginBottom = '200px';
    moreToolsTab.style.marginTop = '-50px';

    // Find a common parent for both the toggle emoji and the tab
    const commonParent = document.querySelector('.margin5'); // Adjust this selector as needed
    commonParent.appendChild(moreToolsTab);






    ///////////////////😎 STYLES CODE ////////////////////
    GM_addStyle(`

    p {
    display: none;   /* Hide/show paragraph text "none" to hide / "block" to show */
}

  #formContainer {
    /* Your styles for the form container */
    position: fixed;
    background-color: rgba(128, 128, 128, 0.7); /* Grey with 70% opacity */
    top: -100%; /* Start hidden above the screen */
    left: 60%; /* Center horizontally */
    transform: translateX(-50%) translateY(0);
    border: 4px solid red;
    padding: 10px;
    border-radius: 5px;
    font-family: 'Courier New', monospace;
    width: 750px;
    box-shadow: 0 0 5px rgba(0, 0, 0, 0.2);
    transition: top 0.3s ease-in-out, opacity 0.3s ease-in-out, max-width 0.3s ease-in-out;
    z-index: 9999;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: space-between;
    opacity: 1; /* Initially visible */
    max-width: 100%; /* Initially visible */
}



#formContainer.active {
    top: 0; /* Slide down from the top when active */
    opacity: 1; /* Show when active */
    max-width: 100%; /* Show when active */
}

.containerTitle {
    font-size: 26px; /* Font size for the container title */
    font-weight: bold;
}




    /* Hover effect */
    #formContainer button:hover {
        background-color: #4CAF50;
    }

    /* Label styles */
    #formContainer > div > div {
        /* Your styles for all labels */
        font-size: 18px;
        Color: white;
        font-family: Arial, sans-serif;
        font-weight: bold;
        margin: 5px;
        text-align: center;
    }

  #customTab {
    position: absolute;
    background-color: transparent; /* Set background color to transparent */
    color: black;
    border-radius: 60px;
    padding: 12px 8px;
    font-size: 22px;
    font-weight: bold;
    cursor: pointer;
    transition: background-color 0.3s, color 0.3s; /* Smooth transition for hover effect */
}

#customTab:hover {
    background-color: transparent;
    padding: 12px 8px;
    font-size: 24px;
    color: black;
}


.divider {
    margin: 0 5px; /* Adjust margin as needed */
    width: 5px; /* Adjust width of the divider */
    height: 5px; /* Adjust height of the divider */
    background-color: #000; /* Color of the divider */
    border-radius: 10%; /* Makes the divider a circle */
    display: inline-block; /* Display the dividers inline */
}




/* BULK EDIT BUTTON */
#part_search_bulk_edit_popup_button {
    border-radius: 20px;
    box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.1);
    background-color:  #fdfefe;
    margin-bottom: 5px;
    font-weight: bold; /* Make text bold */
    transition: background-color 0.3s ease; /* Adding transition for smoother hover and click effect */
}

#part_search_bulk_edit_popup_button:hover {
    background-color:  #dcdcdc; /* Change background color on hover */
}

#part_search_bulk_edit_popup_button:active {
    background-color:  #fdfefe; /* Restore original background color on click */
}






/*MOBILITY SIDE MENU STYLES*/


/* Adjusted animation for slide out effect */
@keyframes slideOut {
  0% {
    transform: translateX(0);
    opacity: 1; /* Start with full opacity */
  }
  25% {
    transform: translateX(-25%);
    opacity: 0.75; /* Reduce opacity a quarter through animation */
  }
  50% {
    transform: translateX(-50%);
    opacity: 0.5; /* Reduce opacity halfway through animation */
  }
  75% {
    transform: translateX(-75%);
    opacity: 0.25; /* Reduce opacity three quarters through animation */
  }
  100% {
    transform: translateX(-100%);
    opacity: 0; /* Fade out as element slides out */
  }
}




.search_refinements {
  font-family: Arial, sans-serif;
  font-size: 16px;
  color: #333;
  background-color: #d0d3d4;
  margin-right: 20px;
  margin-top: -12px;
  padding: 4px; /* Add padding to maintain the grey background */
  border-radius: 10px; /* Add border-radius for curved edges */
  box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2); /* Reduce spread radius for a slight shadow effect */
  border: 2px solid rgba(255, 255, 255, 0.3); /* Set transparent border */
  background-image: linear-gradient(135deg, rgba(255, 255, 255, 0.5), rgba(0, 0, 0, 0.1)); /* Add gradient background */
  transition: transform 0.5s ease; /* Add transition effect for smooth sliding */
}

.row-left7 {
  transition: transform 0.5s ease; /* Add transition effect for smooth sliding */
}

.row-left7.slide {
  transform: translateX(-230px); /* Slide left by 200px */
}




/* Hide the original text */
.search_refinements h4 {
  display: none;
}

/* Add the new text with emoji */
.search_refinements h4::after {
  content: "Refine 🔍"; /* Replace the original text with the new text and emoji */
  font-size: 18px;
  color:  #34495e;
  margin-bottom: 20px; /* Use margin-bottom instead of marginBottom */
  text-decoration: underline; /* Add underline */
  display: inline-block; /* Make the emoji inline with the text */
  margin-left: 5px; /* Adjust margin as needed */
}



/* Headings */
.search_refinements h5 {
  display: inline-block;
  padding: 5px 10px;
  font-size: 14px;
  text-transform: uppercase;
  color: white;
  background-color: #35465c; /* Lighter navy color */
  border-radius: 5px;
  margin-right: 8px;
  margin-bottom: 8px;

  box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2); /* Shadow effect */
  transition: background-color 0.3s; /* Smooth transition for background color */

  text-align: center; /* Center align the text */
}



.search_refinements .linklist {
  list-style-type: none;
  padding-left: 0;
}

.search_refinements .linklist li {
  margin-bottom: 5px;
}


/* DYNAMIC TEXT STYLES */

.search_refinements .linklist li  {
  font-weight: bold; /* Make the dynamic text bold */
  font-size: 14px; /* Adjust the font size as needed */
  color: black; /* Adjust the color as needed */
  text-decoration: none; /* Remove the default underline */
  text-transform: uppercase; /* Uppercase the text */
}

.search_refinements .linklist li :hover {
  text-decoration: underline; /* Add underline on hover if desired */
  color: #007BFF; /* Change color on hover if desired */
}


.search_refinements .linklist li a{
  font-weight: bold; /* Make the dynamic text bold */
  font-size: 13px; /* Adjust the font size as needed */
  color: black; /* Adjust the color as needed */
  text-decoration: none; /* Remove the default underline */
  text-transform: uppercase; /* Uppercase the text */
}

.search_refinements .linklist li a:hover {
  text-decoration: underline; /* Add underline on hover if desired */
  color: #007BFF; /* Change color on hover if desired */
}





/* TABLE COLUMN HEADINGS STYLES*/

.table th a {
    color: white; /* Text color */
    background-color: #31465E; /* Header Bubble Background color */
    padding: 4px; /* Padding */
    text-decoration: none; /* Remove underline */
    border-radius: 5px; /* Rounded corners */
    display: inline-block; /* Make the link a block element */
    width: 100%; /* Expand to full width */
    box-sizing: border-box; /* Include padding and border in element's total width and height */
    text-align: center; /* Center the text */
    margin-bottom: -10px; /* Adjust bottom margin */
    margin-top: -10px; /* Adjust top margin */
    position: relative; /* Position relative for pseudo-elements */
    font-size: 14px; /* Font size */
    font-weight: normal; /* Non-Hyperlinked Header */
    font-family: roboto; /* Font family */
}

/* Hover effect */
.table th a:hover {
    color: white; /* Text color */
    background-color: #35465c; /* Change background color on hover */
}




/* Style for pseudo-elements */
.table th a::before {
    content: ''; /* Empty content */
    position: absolute; /* Position absolutely */
    top: 0; /* Align to the top */
    bottom: 0; /* Align to the bottom */
    left: 0; /* Align to the left */
    right: 0; /* Align to the right */
    background-color: rgba(255, 255, 255, 0.1); /* Semi-transparent white */
    border-radius: 5px; /* Rounded corners */
    transition: opacity 0.3s; /* Smooth transition for opacity */
    pointer-events: none; /* Ignore pointer events */
}

/* Show pseudo-element on hover */
.table th a:hover::before {
    opacity: 1; /* Fully visible on hover */
}




/* STYLE SERIAL ID CELLS */
.table td a {
    color:  #0b36c8; /* Text color */
    text-decoration: none; /* Remove underline */
    font-weight: normal; /* Bold font */
}


/* TABLE HEADER CELL STYLES (Non-hyperlinked) */

.table thead th {
    color: white; /* Text color */
    font-weight: normal;   /*Non-Hyperlinked Header */
    font-family: 'Roboto', sans-serif; /* Font family */
    font-size: 14.5px; /* Font size */
    background-color:      #31465E; /*Header Cell Background color */
    padding-top: 2px; /* Less vertical padding */
    padding-bottom: 2px; /* Less vertical padding */
    padding-left: 4px; /* More horizontal padding */
    padding-right: 4px; /* More horizontal padding */
    text-align: center; /* Center the text horizontally */
    text-transform: uppercase; /* Transform text to uppercase */
    vertical-align: middle; /* Center the text vertically */
    border: 0.2px solid black; /* Thin black border for internal gridlines */
}

/* TABLE HEADER STYLES (hyperlinked) */
.table thead th a {
    /* Your styles here */
    color: white; /* Text color */
    font-weight: normal;   /*Non-Hyperlinked Header */
    font-family: roboto;
    font-size: 14.5px; /* Font size */
    /* Add more styles as needed */
}


/* Style for table data cells */
.table tbody td {
    color: black; /* Text color */
    font-weight: normal; /* Normal font weight for data cells */
    font-size: 13px; /* Font size */
    font-family: Roboto, sans-serif; /* Font family */
    background-color: #f8f9f9; /* Background color */
    padding-top: 4px; /* Less vertical padding */
    padding-bottom: 4px; /* Less vertical padding */
    padding-left: 20px; /* More horizontal padding */
    padding-right: 20px; /* More horizontal padding */
    text-align: center; /* Center the text horizontally */
    vertical-align: middle; /* Center the text vertically */
    border: 0.2px solid black; /* Thin black border for internal gridlines */
    text-transform: uppercase; /* Transform text to uppercase */
}


/* ADVANCED EDIT BUTTON */
#editpane {
    z-index: 9999 !important; /* Give it a very high z-index */
    position: absolute; /* Ensure the position is absolute */
    border-radius: 20px; /* Rounded corners */
    box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.1); /* Add a subtle shadow */
    background-color: #fdfefe; /* Set background color */
    margin-bottom: 5px; /* Margin at the bottom */
    font-weight: bold; /* Make text bold */
    transition: background-color 0.3s ease; /* Smoother hover and click effect */
    padding: 10px; /* Add padding for better spacing */
}


/* ADVANCED EDIT BUTTON */
#Block1 {
    z-index: 9999 !important; /* Give it a very high z-index */
    position: absolute; /* Ensure the position is absolute */
    border-radius: 20px; /* Rounded corners */
    box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.1); /* Add a subtle shadow */
    background-color: #fdfefe; /* Set background color */
    margin-bottom: 5px; /* Margin at the bottom */
    font-weight: bold; /* Make text bold */
    transition: background-color 0.3s ease; /* Smoother hover and click effect */
    padding: 10px; /* Add padding for better spacing */
}




/* Style for the table */
.table {
    border-collapse: collapse; /* Collapse borders to ensure shared borders */
    border-radius: 10px; /* Border radius */
    position: relative; /* Relative positioning for pseudo-element */
}

/* Add glossy and curved border to the outside of the table */
.table:before {
    content: '';
    position: absolute;
    top: -8px; /* Adjust to make the top edge come out more */
    left: -8px; /* Adjust to make the left edge come out more */
    right: -8px; /* Adjust to make the right edge come out more */
    bottom: -8px; /* Adjust to make the bottom edge come out more */
    border-radius: 20px; /* Curved border radius */
    border-top: 8px solid #f5f5f5; /* Lighter color on top */
    border-left: 8px solid #f5f5f5; /* Lighter color on left */
    border-bottom: 8px solid #cccccc; /* Darker color on bottom */
    border-right: 8px solid #cccccc; /* Darker color on right */
    z-index: -1; /* Behind the table */
}



`);






    ////// CONTAINER TITLE ////////////////////////////// TO FIX/IMRPROVE!


    function createTooltip(element, message) {
        const tooltip = document.createElement('span');
        tooltip.classList.add('tooltip-message');
        tooltip.textContent = message;

        element.appendChild(tooltip);

        setTimeout(() => {
            element.removeChild(tooltip);
        }, 2000);
    }




    //////////📈 INCREASE MOBILITY PAGE LIMIT 📈///////////////////

    // Find the select element by ID
    var selectElement = document.getElementById('max_rows');

    // Function to get the value of a URL parameter by name
    function getParameterByName(name, url) {
        if (!url) url = window.location.href;
        name = name.replace(/[\[\]]/g, '\\$&');
        var regex = new RegExp('[?&]' + name + '(=([^&#]*)|&|#|$)'),
            results = regex.exec(url);
        if (!results) return null;
        if (!results[2]) return '';
        return decodeURIComponent(results[2].replace(/\+/g, ' '));
    }

    if (selectElement) {
        // Remove the 'selected' attribute from the default 500 option
        var defaultOption = selectElement.querySelector('option[selected]');
        if (defaultOption) {
            defaultOption.removeAttribute('selected');
        }

        // Create new options for 1000, 2000, and 6000 results per page
        var newOptions = ['1000', '2000', '6000'];

        newOptions.forEach(function(value) {
            var newOption = document.createElement('option');
            newOption.value = value;
            newOption.textContent = value;
            selectElement.appendChild(newOption);
        });

        // Set the dropdown to the value of 'max_rows' in the URL if it exists
        var currentMaxRows = getParameterByName('max_rows');
        if (currentMaxRows) {
            selectElement.value = currentMaxRows;
        }

        // Add event listener to reload the page with the selected value
        selectElement.addEventListener('change', function() {
            var selectedValue = selectElement.value;
            var currentUrl = window.location.href;

            // Update the 'max_rows=' parameter in the URL with the selected value
            var newUrl = currentUrl.replace(/max_rows=\d+/, 'max_rows=' + selectedValue);

            // Reload the page with the new URL
            window.location.href = newUrl;
        });
    }







    ////////📋 COPY BIN NAME 🗑️/////////////////////////

    // Function to copy text to clipboard
    const copyBinToClipboard = (text) => {
        const textarea = document.createElement('textarea');
        textarea.value = text;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
    };

    // Function to handle click event and copy the bin string
    const handleCellClick = (cell, siteIndex, roomIndex, binIndex) => {
        const row = cell.parentElement;
        const siteValue = row.cells[siteIndex].textContent.trim();
        const roomValue = row.cells[roomIndex].textContent.trim();
        const binValue = row.cells[binIndex].textContent.trim();
        const combinedValue = siteValue + '.' + roomValue + '.' + binValue;

        copyBinToClipboard(combinedValue);

        const nextCell = row.cells[binIndex];
        const originalText = nextCell.textContent;
        nextCell.textContent = '🗑️ Copied! ✔️';
        setTimeout(() => {
            nextCell.textContent = originalText;
        }, 500); // Revert back after 1 second
    };

    // Function to find column indexes by heading text
    const getColumnIndexesByHeadings = (table, headings) => {
        const headerCells = table.querySelectorAll('thead th');
        const indexes = {};
        headings.forEach(heading => {
            indexes[heading] = -1; // Default to -1 if not found
        });

        headerCells.forEach((cell, index) => {
            const text = cell.textContent.trim();
            if (headings.includes(text)) {
                indexes[text] = index;
            }
        });

        return indexes;
    };

    // Get the table element
    const table = document.querySelector('.table-bordered.table-striped.table-condensed');

    if (table) {
        // Find the indexes of the "Site," "Room," and "Bin" columns
        const columnIndexes = getColumnIndexesByHeadings(table, ['Site', 'Room', 'Bin']);
        const siteIndex = columnIndexes['Site'];
        const roomIndex = columnIndexes['Room'];
        const binIndex = columnIndexes['Bin'];

        if (siteIndex === -1 || roomIndex === -1 || binIndex === -1) {
            console.error('One or more columns ("Site", "Room", "Bin") not found!');
        } else {
            // Get all rows in the table body
            const rows = document.querySelectorAll('.table-bordered.table-striped.table-condensed tbody tr');

            // Loop through each row and add click events to the "Room" and "Bin" cells
            rows.forEach(row => {
                const cells = row.querySelectorAll('td');

                if (cells.length > binIndex) {
                    const roomCell = cells[roomIndex];
                    const binCell = cells[binIndex];

                    roomCell.style.cursor = 'pointer'; // Change cursor to pointer to indicate clickability
                    roomCell.title = '📋 Click to copy Bin';
                    roomCell.addEventListener('click', () => handleCellClick(roomCell, siteIndex, roomIndex, binIndex));

                    binCell.style.cursor = 'pointer'; // Change cursor to pointer to indicate clickability
                    binCell.title = '📋 Click to copy Bin';
                    binCell.addEventListener('click', () => handleCellClick(binCell, siteIndex, roomIndex, binIndex));
                }
            });
        }
    } else {
        console.error('Table not found!');
    }






    ////////📋 COPY SERIAL ID 🗑️/////////////////////////

    // Function to add a button to a cell in the "Tag" column
    function addButtonToTagCell(cell) {
        // Create the button element
        var button = document.createElement('button');
        button.textContent = 'SN📋';

        // Style the button
        button.style.border = 'none';
        button.style.backgroundColor = 'transparent'; // Set background color to transparent
        button.style.padding = '0'; // Set padding to zero
        button.style.margin = '0'; // Set margin to zero

        // Add a click event handler to the button
        button.addEventListener('click', function() {
            // Get the text from the "Serial ID" column (index 2)
            var serialIdCell = cell.parentElement.cells[2];
            var serialIdText = serialIdCell.textContent.trim();

            // Check if "( RMA )" is present at the end and trim it
            var trimmedText = serialIdText.replace(/\s*\(\s*RMA\s*\)$/i, '');

            // If the text was not trimmed (meaning "( RMA )" was not present), use the original text
            var truncatedText = trimmedText === serialIdText ? serialIdText : trimmedText;

            // Copy the serial ID text to the clipboard
            copyToClipboard2(truncatedText);

            console.log('Serial ID copied:', truncatedText);

            // Change button text to "Copied!" temporarily
            button.textContent = 'SN✔️';

            // Restore button text after 1 second
            setTimeout(function() {
                button.textContent = 'SN📋';
            }, 500);
        });

        // Set padding to 0 for the button's parent element or adjust the padding of its container
        cell.style.padding = '0';

        // Append the button to the cell
        cell.appendChild(button);
    }

    // Function to copy text to the clipboard
    function copyToClipboard2(text) {
        var textarea = document.createElement('textarea');
        textarea.value = text;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
    }

    // Get all rows in the table body
    var rows2 = document.querySelectorAll('.table-bordered.table-striped.table-condensed tbody tr');

    // Loop through each row and add button to the "Tag" column
    rows2.forEach(row => {
        var cells = row.querySelectorAll('td');
        var tagIndex = 3; // Column index for the "Tag" column

        if (cells.length > tagIndex) {
            var cell = cells[tagIndex];
            addButtonToTagCell(cell);
        }
    });



    //////// 📋 COPY CELL DATA //////

    // Function to copy text to clipboard
    const copyModelToClipboard = (text) => {
        const textarea = document.createElement('textarea');
        textarea.value = text;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
    };

    // Function to add click event to specified cells using column header name
    const addCopyEventToCellsByHeader = (headerName) => {
        const table = document.querySelector('table.table-bordered.table-striped.table-condensed');
        if (!table) {
            console.error('Table not found!');
            return;
        }

        // Find the column index by header name
        const headerCells = table.querySelectorAll('thead th');
        let columnIndex = -1;

        headerCells.forEach((cell, index) => {
            if (cell.textContent.trim() === headerName) {
                columnIndex = index + 1; // nth-child uses 1-based index
            }
        });

        if (columnIndex === -1) {
            console.warn(`Column "${headerName}" not found.`);
            return;
        }

        // Add click event to cells in the column
        const cells = table.querySelectorAll(`tbody tr td:nth-child(${columnIndex})`);
        cells.forEach(cell => {
            cell.style.cursor = 'pointer';
            cell.title = `📋 Click to copy ${headerName}`;

            cell.addEventListener('click', () => {
                // Clone the cell to avoid modifying the original content
                const clonedCell = cell.cloneNode(true);
                // Remove any <h4> elements from the clone
                clonedCell.querySelectorAll('h4').forEach(el => el.remove());

                // Extract text from the cleaned clone
                const text = clonedCell.textContent.trim();
                copyModelToClipboard(text);
                console.log(`${headerName} copied:`, text);

                // Temporarily show "Copied!✔️" notification
                const originalHTML = cell.innerHTML;
                cell.textContent = 'Copied!✔️';

                setTimeout(() => {
                    // Restore the original content after 0.5 seconds
                    cell.innerHTML = originalHTML;
                }, 500);
            });
        });
    };

    // Call the function to enable copying for specified headers
    addCopyEventToCellsByHeader('Model');
    addCopyEventToCellsByHeader('Model MPN');
    addCopyEventToCellsByHeader('Model APN');
    addCopyEventToCellsByHeader('Model Description');
    addCopyEventToCellsByHeader('PO Number');
    addCopyEventToCellsByHeader('RMA ID');
    addCopyEventToCellsByHeader('Outbound Tracking ID');
    addCopyEventToCellsByHeader('Tracking ID');
    addCopyEventToCellsByHeader('Outbound RMA ID');
    addCopyEventToCellsByHeader('Build ID');
    addCopyEventToCellsByHeader('Shipment Id');
    addCopyEventToCellsByHeader('Tote Asset ID');
    addCopyEventToCellsByHeader('Type');






    /// TESTING - MPN BIN SEARCH /////

    function modelBinSearch() {
        // Attach an event listener to the table rows
        const table = document.querySelector('.table.table-bordered.table-striped.table-condensed');

        if (!table) {
            console.error('🗑️ BIN SEARCH Table not found!');
            return;
        }

        // Loop through all rows in the table
        const rows = table.querySelectorAll('tbody tr');
        rows.forEach(row => {
            // Get the Site cell (index 8) in the current row
            const siteCell = row.cells[8 - 1]; // Adjust for 0-based index

            if (siteCell) {
                // Add click event listener to the Site cell
                siteCell.addEventListener('click', function() {
                    // Get the Site and MPN values from the row
                    const siteValue = siteCell.textContent.trim();
                    const mpnValue = row.cells[11 - 1]?.textContent.trim(); // Assuming MPN is in the 12th column

                    if (!mpnValue) {
                        console.error('🗑️ BIN SEARCH MPN data not found in the row');
                        return;
                    }

                    // Construct the URL by inserting the Site and MPN values

                    const searchURL = `https://mobility.amazon.com/part/search?search_type=all&search_string=model%3A%22${encodeURIComponent(mpnValue)}%22%20%26%26%20site%3A%22${encodeURIComponent(siteValue)}%22%20%26%26%20state%3A%22SPARE%22&max_rows=6000&query=GO`;

                    console.log(`🗑️ BIN SEARCH Triggered for Site: ${siteValue}, MPN: ${mpnValue}`);
                    console.log(`🗑️ BIN SEARCH Fetching URL: ${searchURL}`);

                    // Perform the fetch request and process the data
                    fetchDataFromURL(searchURL, mpnValue);
                });
            }
        });
    }




    let activePopup = null; // Store the currently open popup
    let activeMPNCell = null; // Store the currently active MPN cell

    function addSiteCellClickListener() {
        const tableRows = document.querySelectorAll('.table.table-bordered.table-striped.table-condensed tbody tr');

        tableRows.forEach((row, rowIndex) => {
            const siteCell = row.cells[7]; // Assuming index 8 for 'Site'
            const mpnCell = row.cells[10]; // Assuming index 12 for 'MPN'
            const stateCell = row.cells[17]; // Assuming index 18 for 'State'

            if (siteCell && mpnCell && stateCell) {
                const siteValue = siteCell.textContent.trim();
                const mpnValue = mpnCell.textContent ? mpnCell.textContent.trim() : "null"; // Handle null or empty MPN
                const stateValue = stateCell.textContent.trim();

                // Add hover title and pointer cursor
                siteCell.title = `Show bins for ${mpnValue}(${stateValue}) at ${siteValue}`;
                siteCell.style.cursor = "pointer";

                // Add click event listener
                siteCell.addEventListener('click', () => {
                    console.log(`🗑️ BIN SEARCH Triggered for Site: ${siteValue}, MPN: ${mpnValue}, State: ${stateValue}`);

                    // === Inject the MPN check here ===
                    // Check if MPN is "null"
                    if (mpnValue === "null" || mpnValue.trim().toLowerCase() === "null") {
                        // Show a quick 2-second popup
                        showQuickPopup("MPN Not Identified!", "red");
                        return; // Stop execution if MPN is not identified
                    }

                    // Close any previously open popup
                    closeActivePopup();
                    // Highlight the current MPN cell
                    if (activeMPNCell && activeMPNCell !== mpnCell) {
                        activeMPNCell.style.backgroundColor = ""; // Reset the previous MPN cell color
                    }

                    // Keep the current MPN cell highlighted in light green
                    mpnCell.style.backgroundColor = "lightgreen";

                    // Ensure activeMPNCell is updated to the currently clicked one
                    activeMPNCell = mpnCell;


                    // Construct the URL and trigger bin search
                    const url = `https://mobility.amazon.com/part/search?search_type=all&search_string=model%3A%22${mpnValue}%22%20%26%26%20site%3A%22${siteValue}%22%20%26%26%20state%3A%22${stateValue}%22&max_rows=6000&query=GO`;

                    // Perform the search and display popup with bins
                    fetchDataFromURL(url, mpnValue, siteValue, stateValue);
                });
            }
        });
    }

    // Function to show a quick popup error message
    function showQuickPopup(message, backgroundColor) {
        // Create a small div for the popup
        const popup = document.createElement('div');
        popup.style.position = 'fixed';
        popup.style.top = '20%';
        popup.style.left = '50%';
        popup.style.transform = 'translate(-50%, -50%)';
        popup.style.width = '200px';
        popup.style.backgroundColor = backgroundColor || '#ff4d4d'; // Red by default
        popup.style.color = 'white';
        popup.style.textAlign = 'center';
        popup.style.padding = '15px';
        popup.style.borderRadius = '8px';
        popup.style.fontSize = '16px';
        popup.style.zIndex = '9999';
        popup.innerText = message;

        // Append the popup to the body
        document.body.appendChild(popup);

        // Remove the popup after 2 seconds
        setTimeout(() => {
            document.body.removeChild(popup);
        }, 2000); // 2 seconds
    }


    // Function to close the active popup
    function closeActivePopup() {
        if (activePopup) {
            document.body.removeChild(activePopup); // Remove the popup from the DOM
            activePopup = null; // Set the active popup to null
        }
        if (activeMPNCell) {
            activeMPNCell.style.backgroundColor = ""; // Reset the MPN cell color after closing
            activeMPNCell = null; // Reset the active MPN cell
        }
    }


    // Fetch data and display the bins in the popup
    async function fetchDataFromURL(url, mpnValue, siteValue, stateValue, binList = null) {
        console.log(`Attempting to fetch data with URL: ${url}`); // Log the fetch attempt
        try {
            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Content-Type': 'text/html',
                },
            });

            if (!response.ok) {
                console.error('🗑️ BIN SEARCH Failed to fetch data:', response.statusText);
                return;
            }

            const htmlText = await response.text();
            console.log('🗑️ BIN SEARCH Fetched HTML response:', htmlText);

            // Parse the HTML and extract bin names and their counts
            const parser = new DOMParser();
            const doc = parser.parseFromString(htmlText, 'text/html');

            const allSpans = doc.querySelectorAll('span');
            const binCounts = {};

            allSpans.forEach((span, index) => {
                if (span.textContent.includes('Site.Room.Bin:')) {
                    const nextSpan = allSpans[index + 1]; // Get the next <span> element
                    if (nextSpan) {
                        const binName = nextSpan.textContent.trim();
                        binCounts[binName] = (binCounts[binName] || 0) + 1;
                    }
                }
            });

            const uniqueBins = Object.keys(binCounts);
            console.log('🗑️ BIN SEARCH Unique bins found:', uniqueBins);

            // If no results, display a message in the bin area
            if (uniqueBins.length === 0) {
                displayNoBinsFoundMessage(siteValue, mpnValue, stateValue, binList);
                return;
            }

            // If binList is passed, update it. Otherwise, create a new popup
            if (binList) {
                updateBinList(binList, uniqueBins, binCounts);
            } else {
                displayBinPopup(uniqueBins, binCounts, mpnValue, url, siteValue, stateValue);
            }
        } catch (error) {
            console.error('🗑️ BIN SEARCH Error fetching or processing data:', error);
        }
    }

    // Display the popup with dropdowns and bin results
    function displayBinPopup(uniqueBins, binCounts, mpnValue, url, siteValue, stateValue) {
        // Close any previously active popup before opening a new one
        closeActivePopup();

        // Create a popup window
        const popup = document.createElement('div');
        popup.style.position = 'fixed';
        popup.style.top = '40%';
        popup.style.left = '30%';
        popup.style.transform = 'translate(-50%, -50%)';
        popup.style.width = '500px';
        popup.style.backgroundColor = 'rgba(244, 208, 63, 0.85)'; // Light orange background with 85% opacity
        popup.style.border = '2px solid black';
        popup.style.boxShadow = '0 8px 15px rgba(0,0,0,0.4)'; // 3D effect
        popup.style.padding = '20px';
        popup.style.zIndex = '9999';
        popup.style.borderRadius = '12px';

        activePopup = popup; // Set this popup as the active one

        // Create the centered title for the popup
        const popupTitle = document.createElement('h2');
        popupTitle.textContent = "Bin Lookup 🔎🗑️";
        popupTitle.style.fontSize = '22px';
        popupTitle.style.color = 'black';
        popupTitle.style.textAlign = 'center';  
        popupTitle.style.marginBottom = '1px';
        popupTitle.style.backgroundColor = 'lightgrey';  
        popupTitle.style.borderRadius = '12px';  
        popupTitle.style.padding = '2px';  
        popupTitle.style.width = '50%';  
        popupTitle.style.margin = '0 auto';  
        popupTitle.style.boxSizing = 'border-box';  
        popupTitle.style.border = '1px solid black';  
        popup.appendChild(popupTitle);




        // Title displaying site, mpn, and state
        const title = document.createElement('h3');
        title.innerHTML = `📍 SITE: <b>${siteValue}</b><br>🔢 MODEL: <b>${mpnValue}</b> (${stateValue})`;
        title.style.fontSize = '16px';
        title.style.color = 'black';
        title.style.textAlign = 'left';
        title.style.marginBottom = '15px';

        // Create the link emoji
        const linkEmoji = document.createElement('a');
        linkEmoji.href = url;
        linkEmoji.textContent = '🔗';
        linkEmoji.style.fontSize = '20px';
        linkEmoji.style.marginLeft = '10px'; // Add a little space before the emoji
        linkEmoji.target = '_blank'; // Open link in a new tab

        // Append the link emoji to the title
        title.appendChild(linkEmoji);

        // Append the title (with the link) to the popup
        popup.appendChild(title);

        // Create the dropdown for siteValue
        const siteDropdown = document.createElement('select');
        const siteOptions = [
            "AKL55", "AKL60", "AKL61", "AKL70", "ALE60", "ALE61", "ALE62", "APA60", "APA61", "APA62",
            "ARN50", "ARN51", "ARN52", "ARN57", "ATL700", "AYQ60", "BAH52", "BAH53", "BAH54", "BJS11",
            "BJS12", "BJS20", "BJS50", "BJS60", "BJS70", "BJS71", "BJS80", "BOM52", "BOM53", "BOM54",
            "BOM56", "BOM59", "BOM60", "BOM61", "BOM62", "BOM63", "BOM64", "BOM66", "BPM54", "BPM60",
            "BPM61", "BPM62", "BWI700", "CDG53", "CDG54", "CDG55", "CDG63", "CDG69", "CDG70", "CDG83",
            "CDG84", "CDG93", "CGK60", "CGK61", "CGK62", "CGK63", "CGK64", "CGK65", "CJJ700", "CMH50",
            "CMH51", "CMH52", "CMH53", "CMH54", "CMH55", "CMH56", "CMH57", "CMH58", "CMH59", "CMH60",
            "CMH600", "CMH601", "CMH602", "CMH61", "CMH62", "CMH63", "CMH64", "CMH70", "CMH72", "CMH82",
            "CMH86", "CMH91", "CMH95", "CPT60", "CPT61", "CPT62", "DCA50", "DCA51", "DCA52", "DCA53",
            "DCA54", "DCA56", "DCA57", "DCA58", "DCA60", "DCA62", "DCA64", "DEN22", "DEN700", "DFW700",
            "DUB10", "DUB104", "DUB2", "DUB32", "DUB4", "DUB50", "DUB51", "DUB52", "DUB53", "DUB54",
            "DUB55", "DUB56", "DUB57", "DUB58", "DUB60", "DUB62", "DUB64", "DUB66", "DUB68", "DUB69",
            "DUB74", "DUB78", "DUB8", "DUB84", "DUB9", "DUB91", "DUB94", "DUB99", "DXB60", "DXB61",
            "DXB62", "FRA40", "FRA41", "FRA52", "FRA53", "FRA54", "FRA56", "FRA60", "FRA61", "FRA62",
            "FRA63", "FRA72", "FRA74", "FRA82", "FRA90", "GRU1", "GRU2", "GRU4", "GRU50", "GRU51",
            "GRU60", "GRU61", "GRU62", "GRU63", "GRU65", "GRU67", "HEF60", "HEF61", "HEF62", "HKG60",
            "HKG61", "HKG62", "HKG63", "HYD100", "HYD101", "HYD110", "HYD111", "HYD112", "HYD120",
            "HYD121", "HYD122", "HYD60", "IAD1", "IAD10", "IAD100", "IAD101", "IAD102", "IAD103",
            "IAD104", "IAD108", "IAD109", "IAD11", "IAD111", "IAD114", "IAD115", "IAD116", "IAD117",
            "IAD118", "IAD119", "IAD12", "IAD120", "IAD121", "IAD122", "IAD129", "IAD13", "IAD130",
            "IAD131", "IAD132", "IAD133", "IAD134", "IAD135", "IAD136", "IAD14", "IAD140", "IAD141",
            "IAD142", "IAD143", "IAD144", "IAD145", "IAD146", "IAD147", "IAD148", "IAD149", "IAD15",
            "IAD150", "IAD151", "IAD152", "IAD155", "IAD156", "IAD157", "IAD158", "IAD159", "IAD16",
            "IAD161", "IAD162", "IAD163", "IAD164", "IAD165", "IAD175", "IAD180", "IAD24", "IAD32",
            "IAD35", "IAD50", "IAD51", "IAD52", "IAD54", "IAD55", "IAD56", "IAD57", "IAD58", "IAD59",
            "IAD6", "IAD60", "IAD600", "IAD601", "IAD602", "IAD603", "IAD604", "IAD606", "IAD607",
            "IAD608", "IAD609", "IAD61", "IAD614", "IAD62", "IAD63", "IAD64", "IAD65", "IAD68", "IAD7",
            "IAD71", "IAD73", "IAD74", "IAD75", "IAD76", "IAD77", "IAD78", "IAD79", "IAD80", "IAD81",
            "IAD83", "IAD84", "IAD85", "IAD86", "IAD88", "IAD89", "IAD9", "IAD90", "IAD91", "IAD92",
            "IAD93", "IAD95", "IAD96", "IAD98", "IAD99", "ICN400", "ICN52", "ICN53", "ICN54", "ICN56",
            "ICN63", "ICN65", "ICN66", "ICN80", "KIX50", "KIX51", "KIX52", "KIX53", "KIX54", "KIX55",
            "KIX56", "KIX60", "KIX61", "KIX700", "KUL60", "KUL61", "KUL62", "KUL63", "LAS53", "LAS700",
            "LAX61", "LAX62", "LCK50", "LCK51", "LCK52", "LCK53", "LCK54", "LCK55", "LCK60", "LCK61",
            "LHR53", "LHR54", "LHR55", "LHR56", "LHR57", "LHR59", "LHR64", "LHR65", "LHR74", "LHR79",
            "LHR82", "LHR99", "LTW60", "LTW61", "LTW62", "LUX2", "LUX8", "LUX9", "MAD55", "MCI60",
            "MCI61", "MEL60", "MEL61", "MEL62", "MIA700", "MUC350", "MXP60", "MXP61", "MXP62", "MXP63",
            "MXP73", "MXP74", "NCL60", "NCL61", "NCL62", "NRT11", "NRT12", "NRT20", "NRT22", "NRT55",
            "NRT56", "NRT6", "NRT62", "NRT64", "NRT66", "NRT67", "NRT69", "NRT7", "NRT700", "NRT72",
            "NRT78", "NRT8", "OSU60", "OSU61", "OSU62", "OSU63", "OSU64", "OSU65", "PDT1", "PDT109",
            "PDT130", "PDT2", "PDT4", "PDT50", "PDT51", "PDT53", "PDT54", "PDT55", "PDT56", "PDT58",
            "PDT59", "PDT60", "PDT61", "PDT62", "PDT63", "PDT64", "PDT65", "PDT66", "PDT67", "PDT90",
            "PDT91", "PDX1", "PDX109", "PDX110", "PDX130", "PDX138", "PDX146", "PDX154", "PDX162",
            "PDX170", "PDX178", "PDX2", "PDX4", "PDX50", "PDX51", "PDX52", "PDX53", "PDX54", "PDX55",
            "PDX56", "PDX57", "PDX58", "PDX59", "PDX60", "PDX61", "PDX62", "PDX63", "PDX64", "PDX65",
            "PDX66", "PDX67", "PDX68", "PDX69", "PDX80", "PDX81", "PDX82", "PDX84", "PDX90", "PDX91",
            "PDX92", "PDX93", "PDX95", "PDX96", "PDX97", "PDX98", "PDX99", "PEK50", "PEK7", "PHX53", "PHX52", "PHX100", "PHX50", "PUQ60",
            "QRO60", "QRO61", "QRO62", "SEA14", "SEA15", "SEA3", "SEA300", "SEA31", "SEA32", "SEA700",
            "SEA72", "SFO11", "SFO20", "SFO21", "SFO53", "SFO6", "SFO60", "SFO61", "SFO69", "SFO7",
            "SFO8", "SFO9", "SIN2", "SIN3", "SIN4", "SIN50", "SIN51", "SIN53", "SIN54", "SIN58", "SIN60",
            "SIN61", "SIN63", "SIN64", "SIN65", "SIN67", "SMF700", "SYD5", "SYD51", "SYD52", "SYD53",
            "SYD54", "SYD55", "SYD56", "SYD6", "SYD61", "SYD62", "SYD64", "SYD66", "SYD7", "SYD70",
            "SYD71", "TLV60", "TLV61", "TLV62", "TTN700", "YUL50", "YUL51", "YUL52", "YUL54", "YUL55",
            "YUL63", "YUL71", "ZAZ60", "ZAZ61", "ZAZ62", "ZHY50", "ZHY51", "ZHY52", "ZRH60", "ZRH61",
            "ZRH62"
        ];

        // Set the selected option based on the current siteValue
        siteOptions.forEach(site => {
            const siteOption = document.createElement('option');
            siteOption.value = site;
            siteOption.text = site;
            if (site === siteValue) {
                siteOption.selected = true; // Set this option as selected
            }
            siteDropdown.appendChild(siteOption);
        });


        // Set a fixed width for the site dropdown
        siteDropdown.style.width = '120px'; // Adjust as necessary

        // Create the dropdown for stateValue with all the provided options
        const stateDropdown = document.createElement('select');
        const stateOptions = [
            "BROKEN", "BUILD_SPARE", "BUILD_SPARE_HOLDING", "CONSUMED", "DEGAUSSED", "DEPLOYED", "DESTROYED", "FAILED", "FAILED_IN_USER_CUSTODY",
            "FAILED_RECEIVE", "FAILED_RECEIVE_IN_USER_CUSTODY", "FOUND", "IN_TRANSIT", "MISSING", "OMD", "OUT_FOR_BUILD", "OUT_FOR_DEPLOY", "OUT_FOR_REPAIR", "OUT_OF_DC",
            "PENDING_BUILD", "PENDING_DEGAUSS", "PENDING_DEGAUSS_IN_USER_CUSTODY", "PENDING_DESTRUCTION", "PENDING_DESTRUCTION_IN_USER_CUSTODY", "PENDING_HUB", "PENDING_OMD", "PENDING_RECYCLING", "PENDING_RESALE", "PENDING_REUSE",
            "PENDING_RMA", "PENDING_RMA_IN_USER_CUSTODY", "PENDING_SANITIZATION", "PENDING_SANITIZATION_FOR_RMA", "PENDING_SANITIZATION_FOR_RMA_IN_USER_CUSTODY", "PENDING_SANITIZATION_IDENTIFICATION", "PENDING_SANITIZATION_IDENTIFICATION_IN_USER_CUSTODY", "PENDING_SANITIZATION_IN_USER_CUSTODY", "PENDING_TESTING", "PENDING_TRANSFER",
            "PENDING_WINSTON_WOLFE", "RECEIVED", "RECOVERED_GEAR", "RECYCLING", "RESERVED_FOR_BUILD", "RESERVED_FOR_REPAIR", "RESALE", "REUSE", "RMA", "RMA_PROCESSED",
            "RMA_PROCESSED_IN_USER_CUSTODY", "SANITIZED", "SANITIZED_FOR_RMA_IN_USER_CUSTODY", "SANITIZED_IN_USER_CUSTODY", "SANITIZED_LIQUIDATED", "SANITIZED_SHREDDED", "SANITIZED_SHREDDED_PENDING_RMA", "SANITIZED_SHREDDED_RMA", "SPARE", "TESTED",
            "TRANSFER_BUILD", "TRANSFER_BUILD_SPARE", "TRANSFER_BUILD_SPARE_HOLDING", "TRANSFERRED"

        ];

        stateOptions.forEach(state => {
            const option = document.createElement('option');
            option.value = state;
            option.text = state;
            stateDropdown.appendChild(option);
        });

        // Set the initial selected state value
        stateDropdown.value = stateValue;

        // Set a fixed width for the state dropdown
        stateDropdown.style.width = '150px'; // Adjust as necessary

        // Add event listeners to dropdowns to trigger a search on change
        siteDropdown.addEventListener('change', () => {
            const newSiteValue = siteDropdown.value;
            updateBins(newSiteValue, mpnValue, stateDropdown.value, binList, linkEmoji);
        });

        // Event listener for state dropdown change
        stateDropdown.addEventListener('change', () => {
            const newStateValue = stateDropdown.value;

            // Update title with new stateValue, but without overwriting the linkEmoji
            title.innerHTML = `📍 SITE: <b>${siteDropdown.value}</b><br>🔢 MODEL: <b>${mpnValue}</b> (${newStateValue})`;

            // Re-append the linkEmoji to the title
            title.appendChild(linkEmoji);

            // Perform a new search with updated state value
            updateBins(siteDropdown.value, mpnValue, newStateValue, binList, linkEmoji);
        });

        // Event listener for site dropdown change
        siteDropdown.addEventListener('change', () => {
            const newSiteValue = siteDropdown.value;

            // Update title with new siteValue, but without overwriting the linkEmoji
            title.innerHTML = `📍 SITE: <b>${newSiteValue}</b><br>🔢 MODEL: <b>${mpnValue}</b> (${stateDropdown.value})`;

            // Re-append the linkEmoji to the title
            title.appendChild(linkEmoji);

            // Perform a new search with updated site value
            updateBins(newSiteValue, mpnValue, stateDropdown.value, binList, linkEmoji);
        });




        // DROPDOWNS APPEND CODE
        const dropdownContainer = document.createElement('div');

        // Ensure dropdowns are aligned on the same line
        dropdownContainer.style.display = 'flex';
        dropdownContainer.style.alignItems = 'center'; // Align the dropdowns vertically center
        dropdownContainer.style.gap = '5px'; // Optional: Adds a small gap between the dropdowns

        // Adjust dropdowns to have consistent width and left alignment
        siteDropdown.style.width = '90px'; // Adjust as necessary
        stateDropdown.style.width = '130px'; // Adjust as necessary

        // Remove margins and set to inline-block to avoid gaps
        siteDropdown.style.marginRight = '0';
        stateDropdown.style.marginLeft = '0';

        // Append dropdowns to the container
        dropdownContainer.appendChild(siteDropdown);
        dropdownContainer.appendChild(stateDropdown);

        // Append the container after the title in the popup
        popup.appendChild(dropdownContainer);



        const binList = document.createElement('ul');
        binList.style.padding = '0';// Ensure no padding on the container
        binList.style.margin = '0'; // Ensure no margin on the container

        uniqueBins.forEach(bin => {
            const listItem = document.createElement('li');
            listItem.style.textAlign = 'left';// Align list item content to the left


            // Create a span element for the bin name
            const binName = document.createElement('span');
            binName.textContent = `🗑️ ${bin} (${binCounts[bin]})`;
            binName.style.cursor = 'pointer';
            binName.style.color = 'blue';
            binName.style.fontSize = '18px'; // Make text bigger
            binName.style.textDecoration = 'underline';
            binName.style.textAlign = 'left'; // Ensure alignment to the left
            binName.style.display = 'block'; // Block display for full width
            binName.style.marginBottom = '10px'; // Space between items



            // Add click-to-copy functionality for the bin name
            binName.addEventListener('click', function() {
                navigator.clipboard.writeText(bin).then(() => {
                    console.log(`🗑️ BIN SEARCH Bin ${bin} copied to clipboard`);

                    // Show green tick after copying
                    const tick = document.createElement('span');
                    tick.textContent = '✔️';
                    tick.fontSize = '18px';
                    tick.style.color = 'green';
                    binName.appendChild(tick);

                    // Remove the tick after 2 seconds
                    setTimeout(() => {
                        binName.removeChild(tick);
                    }, 500);
                });
            });

            listItem.appendChild(binName);
            binList.appendChild(listItem);
        });

        popup.appendChild(binList);

        // Create a close button (Red X at the top right)
        const closeButton = document.createElement('span');
        closeButton.innerHTML = '❌';
        closeButton.style.position = 'absolute';
        closeButton.style.top = '10px';
        closeButton.style.right = '10px';
        closeButton.style.cursor = 'pointer';
        closeButton.style.fontSize = '20px';
        closeButton.style.color = 'red';

        // Close popup on click
        closeButton.addEventListener('click', function() {
            closeActivePopup(); // Close the popup
        });

        popup.appendChild(closeButton);
        document.body.appendChild(popup);

        // Add event listener to close popup when clicking outside
        document.addEventListener('click', function(event) {
            if (activePopup && !popup.contains(event.target)) {
                closeActivePopup();
            }
        }, { once: true });
    }

    // Function to update bins based on dropdown changes
    function updateBins(newSiteValue, mpnValue, newStateValue, binList, linkEmoji) {
        const newUrl = `https://mobility.amazon.com/part/search?search_type=all&search_string=model%3A%22${mpnValue}%22%20%26%26%20site%3A%22${newSiteValue}%22%20%26%26%20state%3A%22${newStateValue}%22&max_rows=6000&query=GO`;
        console.log(`Dropdown changed. Fetching new data with URL: ${newUrl}`);

        // Update the link emoji with the new URL
        linkEmoji.href = newUrl;

        // Perform the fetch request to get the actual bins data
        fetch(newUrl)
            .then(response => response.text())
            .then(htmlText => {
            // Parse the fetched HTML response
            const parser = new DOMParser();
            const doc = parser.parseFromString(htmlText, 'text/html');

            // Extract bin data from the parsed HTML
            const allSpans = doc.querySelectorAll('span');
            const binCounts = {};

            allSpans.forEach((span, index) => {
                if (span.textContent.includes('Site.Room.Bin:')) {
                    const nextSpan = allSpans[index + 1]; // Get the next <span> element
                    if (nextSpan) {
                        const binName = nextSpan.textContent.trim();
                        binCounts[binName] = (binCounts[binName] || 0) + 1;
                    }
                }
            });

            const uniqueBins = Object.keys(binCounts);

            // Clear the existing binList
            binList.innerHTML = '';

            // If no bins are found, display a message
            if (uniqueBins.length === 0) {
                const noBinsMessage = document.createElement('li');
                noBinsMessage.textContent = 'No bins found for the selected criteria';
                noBinsMessage.style.color = 'red';
                noBinsMessage.style.textAlign = 'center';
                binList.appendChild(noBinsMessage);
                return;
            }

            // Update the bin list with actual bins
            uniqueBins.forEach(bin => {
                const listItem = document.createElement('li');
                listItem.style.textAlign = 'left';  // Align list item content to the left

                // Create a span element for the bin name
                const binName = document.createElement('span');
                binName.textContent = `🗑️ ${bin} (${binCounts[bin]})`;
                binName.style.cursor = 'pointer';
                binName.style.color = 'blue';
                binName.style.fontSize = '18px';  // Make text bigger
                binName.style.textDecoration = 'underline';
                binName.style.textAlign = 'left';  // Ensure alignment to the left
                binName.style.display = 'block';  // Block display for full width
                binName.style.marginBottom = '10px';  // Space between items

                // Add click-to-copy functionality for the bin name
                binName.addEventListener('click', function() {
                    navigator.clipboard.writeText(bin).then(() => {
                        console.log(`🗑️ BIN SEARCH Bin ${bin} copied to clipboard`);

                        // Show green tick after copying
                        const tick = document.createElement('span');
                        tick.textContent = '✔️';
                        tick.style.fontSize = '18px';
                        tick.style.color = 'green';
                        binName.appendChild(tick);

                        // Remove the tick after 2 seconds
                        setTimeout(() => {
                            binName.removeChild(tick);
                        }, 500);
                    });
                });

                listItem.appendChild(binName);
                binList.appendChild(listItem);
            });
        })
            .catch(error => {
            console.error('🗑️ BIN SEARCH Error fetching bins:', error);
        });
    }



    // Display "No bins found" message in the bin area
    function displayNoBinsFoundMessage(siteValue, mpnValue, stateValue, binList) {
        // Clear the binList and show the message
        binList.innerHTML = `<li style="color:red; font-size:18px; text-align:center;">No bins found for Site: <b>${siteValue}</b> and MODEL: <b>${mpnValue}</b> (${stateValue})</li>`;
    }



    // Call the function to add click listeners to the site cells
    addSiteCellClickListener();












    ////////////📦📝🚚 Request RMA COLLECTION  ////////////////////////////



    // Specify the index of the Part Id column (0-based)
    const partIdColumnIndex = 1; // Adjust this index to match the correct column

    // Column indices for the required data
    const ClusterIndex = 6;
    const serialIdIndex = 2; // Adjust this index to match the correct column
    const siteIndex = 7; // Adjust these indices to match the correct columns
    const modelMpnIndex = 11;
    const vendorIndex = 15;
    const typeIndex = 16;
    const outboundRmaIndex = 26;


    //VENDOR-CONDITIONAL CC EMAILS
    const vendorEmailMap = {
        'Juniper': {
            cc: ['juniperar@onprocess.com','zahrav@juniper.net', 'support@juniper.net' ],
            mainRecipient: ['asset-recovery@juniper.net']
        },
        'Ciena': {
            cc: ['vendorCC@ciena.com'],
            mainRecipient: 'AssetRecovery@ciena.com'
        },
        // ... (other vendors)
    };





    //SITE-CONDITIONAL CC EMAILS
    const siteEmailMap = {
        'BAH': {
            cc: ['bah-logistics@amazon.com']
        },
        'DXB': {
            cc: ['dxb-logistics@amazon.com']
        },
        'ARN': {
            cc: ['arn-logistics@amazon.com']
        },
        'FRA': {
            cc: ['fra-logistics@amazon.com']
        },
        'LHR': {
            cc: ['lhr-logistics@amazon.com']
        },

        'CPT': {
            cc: ['cpt-logistics@amazon.com']
        },
        'DUB': {
            cc: ['dub-logistics@amazon.com']
        },
        'CDG': {
            cc: ['cdg-logistics@amazon.com']
        },
        'MXP': {
            cc: ['mxp-logistics@amazon.com']
        },
        'TLV': {
            cc: ['tlv-logistics@amazon.com']
        },
        'ZAZ': {
            cc: ['zaz-logistics@amazon.com']
        },
        'ZRH': {
            cc: ['zrh-logistics@amazon.com']
        },

        // Add more prefix-specific entries as needed
    };






    // MAP MPN to WEIGHT and DIMENSIONS
    const modelMpnData = {
        'PTX1000-72Q-CHAS-S': {
            weight: '31 kg',
            dimensions: '44.2 x 8.8 x 78.7 cm'
        },
        'MPN456': {
            weight: '8 kg',
            dimensions: '40x25x18 cm'
        },
        'NFX250-S2': {
            weight: '4.3 kg',
            dimensions: '44.09 x 4.37 x 30.48 cm'
        },
        'PTX1000-72Q-AC': {
            weight: '30 kg',
            dimensions: '44.2 x 8.8 x 78.7 cm'
        },
        'SRX1500-AC': {
            weight: '7.30 kg',
            dimensions: '43.9 x 4.44 x 46.22 cm'
        },
        'QFX5000-35-JAS': {
            weight: '10.8 kg',
            dimensions: '44.09 x 4.37 x 52.02 cm'
        },
        'EX2300-C-12P': {
            weight: '3.17 kg',
            dimensions: '27.9 x 4.4 x 23.9 cm'
        },
        'EX4300-48P': {
            weight: '7.3 kg',
            dimensions: '44.1 x 4.37 x 46.7 cm'
        },
        'QFX5100-24Q-3AFO': {
            weight: '9.8 kg',
            dimensions: '44.09 x 4.37 x 54.54 cm'
        },
        'QFX5100-24Q-AFO': {
            weight: '9.8 kg',
            dimensions: '44.09 x 4.37 x 54.54 cm'
        },
        'QFX5100-48S-3AFO': {
            weight: '9.9 kg',
            dimensions: '44.09 x 4.37 x 54.54 cm'
        },
        'QFX5100-48S-AFO': {
            weight: '9.9 kg',
            dimensions: '44.09 x 4.37 x 54.54 cm'
        },
        'QFX5100-48T-AFO': {
            weight: '11.2 kg',
            dimensions: '44.09 x 4.37 x 54.54 cm'
        },
        'QFX5100-96S-AFO': {
            weight: '14.74 kg',
            dimensions: '44.09 x 8.8 x 57 cm'
        },
        'QFX5110-48S-AFO': {
            weight: '11 kg',
            dimensions: '44.09 x 4.37 x 52.02 cm'
        },
        'QFX5200-32C-AFO': {
            weight: '11 kg',
            dimensions: '44.09 x 4.31 x 51.5 cm'
        },
        'SRX5K-RE-1800X4': {
            weight: '1.1 kg',
            dimensions: '44.09 x 4.31 x 51.5 cm'
        },
        // Add more entries for other MPNs
    };







    // GET ALL the ROWS in the TABLE BODY
    const rows1 = document.querySelectorAll('.table-bordered.table-striped.table-condensed tbody tr');

    // Define recipientEmail and allCcEmails at a global scope
    const recipientEmail = 'defaultRecipient@example.com';
    const allCcEmails = '';


    // LOOP through EACH ROW and ADD BUTTON to the PART ID column
    rows1.forEach(row => {
        const cells = row.querySelectorAll('td');

        if (cells.length > partIdColumnIndex) {
            const partIdCell = cells[partIdColumnIndex];





            /////🚚 COLLECTION BUTTON STYLES/////////////////////////////////////////////////////

            // Create the button element
            const RMACollectionButton = document.createElement('button');
            RMACollectionButton.textContent = '📦';

            // Add styles to the button
            RMACollectionButton.style.backgroundColor = 'transparent'; // Set background color to transparent
            RMACollectionButton.style.color = 'lightgrey';
            RMACollectionButton.style.border = 'none'; // Remove border
            RMACollectionButton.style.cursor = 'pointer';
            RMACollectionButton.style.fontSize = '18px'; // Add font size here
            RMACollectionButton.style.marginTop = '4px';



            // Set the title attribute for native hover message
            RMACollectionButton.title = '⚠️ Create RMA Collection Request Form / Email';

            // Add hover effect
            RMACollectionButton.addEventListener('mouseenter', function() {
                RMACollectionButton.style.transform = 'scale(1.2)';
            });

            // Remove hover effect
            RMACollectionButton.addEventListener('mouseleave', function() {
                RMACollectionButton.style.transform = 'scale(1)';
            });


            // Append the button to the document body
            document.body.appendChild(RMACollectionButton);


            // Add a CLICK EVENT LISTENER to the button
            RMACollectionButton.addEventListener('click', () => {






                // GET DATA from the SPECIFIC COLUMNS
                let clusterId = cells[ClusterIndex].textContent.trim();

                // Replace clusterId with corresponding text based on conditions
                if (clusterId.includes("BAH")) {
                    clusterId = "Bahrain";
                } else if (clusterId.includes("DUB")) {
                    clusterId = "Dublin";
                } else if (clusterId.includes("LHR")) {
                    clusterId = "London";
                } else if (clusterId.includes("ARN")) {
                    clusterId = "Stockholm";
                } else if (clusterId.includes("TLV")) {
                    clusterId = "Tel Aviv";
                } else if (clusterId.includes("FRA")) {
                    clusterId = "Frankfurt";
                } else if (clusterId.includes("CDG")) {
                    clusterId = "Paris";
                } else if (clusterId.includes("DXB")) {
                    clusterId = "Dubai";
                } else if (clusterId.includes("CPT")) {
                    clusterId = "Cape Town";
                } else if (clusterId.includes("ZRH")) {
                    clusterId = "Zurich";
                } else if (clusterId.includes("ZAZ")) {
                    clusterId = "Zaragoza";
                } // Add more conditions as needed for other airport codes and their corresponding cities

                // Now clusterId contains the desired text based on conditions


                const serialId = cells[serialIdIndex].textContent.trim();
                const site = cells[siteIndex].textContent.trim();
                const modelMpn = cells[modelMpnIndex].textContent.trim();
                const vendor = cells[vendorIndex].textContent.trim();
                const type = cells[typeIndex].textContent.trim();
                const outboundRma = cells[outboundRmaIndex].textContent.trim();
                // GET the CURRENT DATE in DD MONTH YYYY format
                const currentDate = getCurrentDate();


                // Retrieve WEIGHT + DIMENSIONS based on the MPN
                const modelData = modelMpnData[modelMpn] || { weight: 'N/A', dimensions: 'N/A' };
                const weight = modelData.weight;
                const dimensions = modelData.dimensions;

                // Get CURRENT DATE
                function getCurrentDate() {
                    const currentDate = new Date();
                    const day = currentDate.getDate();
                    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
                    const month = monthNames[currentDate.getMonth()];
                    const year = currentDate.getFullYear();
                    return `${day} ${month} ${year}`;
                }


                // GET the EMAIL based on VENDOR
                const vendorInfo = vendorEmailMap[vendor] || {};
                const recipientEmail = vendorInfo.mainRecipient || 'defaultRecipient@example.com';
                const vendorCcEmails = vendorInfo.cc || [];

                // Get SITE-specific CC emails based on the first 3 letters of the site
                const sitePrefix = site.substring(0, 3);
                const siteSpecificCcEmails = siteEmailMap[sitePrefix]?.cc || [];

                // Combine all CC emails
                const allCcEmails = [...vendorCcEmails, ...siteSpecificCcEmails].join('; ');




                // Assuming serialId contains "RMA" and "ABC" and needs cleaning
                const charactersToRemove = 7; // Adjust the number as needed
                const cleanedSerialId = serialId.slice(0, -charactersToRemove).trim();









                /////////    CONTACT INFO FOR EACH SITE    //////////////
                const contactInfo = {
                    //BAH INFO
                    'BAH52': {
                        name: 'PJ Byrne',
                        mobile: '+973-3510-1356',
                        email: 'pjbyrne@amazon.com'
                    },
                    'BAH53': {
                        name: 'Jubish Kannoly',
                        mobile: '+973-3949-0916',
                        email: ' jubish@amazon.com'
                    },
                    'BAH54': {
                        name: 'PJ Byrne',
                        mobile: '+973-3510-1356',
                        email: 'pjbyrne@amazon.com'
                    },



                    'Other Sites': {
                        name: '[INSERT NAME]',
                        mobile: '[MOBILE]',
                        email: '[EMAIL]'
                    }
                };

                // Get the contact information based on the selected site
                const selectedContact = contactInfo[site] || contactInfo['Other Sites'];

                // Convert TYPE and VENDOR to uppercase
                const uppercaseType = type.toUpperCase();
                const uppercaseVendor = vendor.toUpperCase();




                ////////////////////EMAIL BODY//////////////////////////////////////////////


                // Generate the CONTACT TABLE
                const contactTable = `CONTACT
╔═══════════════════════════════════╤═══════════════════════════════════════╗
║ NAME                              MOBILE                                               EMAIL
╟═══════════════════════════════════╪═══════════════════════════════════════╢
║ ${selectedContact.name}             ${selectedContact.mobile}                             ${selectedContact.email}
╚═══════════════════════════════════╧═══════════════════════════════════════╝`;

                // Determine the value for FREE TRADE ZONE
                const freeTradeZone = site === "DXB60" ? "YES" : "NO";



                // Customize the email subject and body
                const subject = `[ ${site} ] RMA Pickup Request  ║  ${uppercaseVendor}  ║  RMA#: ${outboundRma}   SN: ${cleanedSerialId}  ║  ${currentDate}`;


                const body = `Hello ${uppercaseVendor} Team,

Below part is ready for collection. Please arrange collection (or provide AirwayBill if applicable), thank you:
....................................................................................
╔════════════════════════════════════════╗
║               COLLECTION REQUEST FORM              ║
╔═══════════════════╤════════════════════╗
║ [SN]                                 │ ${cleanedSerialId}
╟═══════════════════╪════════════════════╢
║ [MPN]                             │ ${modelMpn}
╟═══════════════════╪════════════════════╢
║ [TYPE]                              │ ${uppercaseType}
╟═══════════════════╪════════════════════╢
║ [RMA#]                           │ ${outboundRma}
╟═══════════════════╪════════════════════╢
║ [SITE]                               │ ${site}
╟═══════════════════╪════════════════════╢
║ [WEIGHT]                       │ ${weight}
╟═══════════════════╪════════════════════╢
║ [DIM]                               │ ${dimensions}
╟═══════════════════╪════════════════════╢
║ [PALLETIZED?]                │     NO
╟═══════════════════╪════════════════════╢
║ [FREE TRADE ZONE?]    │    ${freeTradeZone}
╚═══════════════════╧════════════════════╝

${contactTable}

Thank you.

Amazon ${clusterId} Team
`;


                // CREATE the MAIL-TO link
                const mailtoLink = `mailto:${recipientEmail}?cc=${encodeURIComponent(allCcEmails)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

                // SET the LOCATION to the MAIL-TO link
                window.location.href = mailtoLink;
            });

            // APPEND the BUTTON to the PART ID cell
            partIdCell.appendChild(RMACollectionButton);
        }
    });














    ////////////////////🏗️🧩✍🏼 CREATE BUTTONS/DROPDOWNS/TEXTBOXES CODE //////////////////////////////////





    //////////////// 💬 REQUEST SAN MESSAGE BUTTON  //////////////////////     *23s to 7s


    function copyDataToClipboard() {
        const targetColumnsIndices = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]; // Update with the required column indices
        const tableRows = document.querySelectorAll('table.table-bordered.table-striped.table-condensed tbody tr');
        if (!tableRows.length) {
            console.error('Table rows not found');
            return;
        }

        // Function to extract text content while excluding specific HTML elements
        function extractTextContent(element) {
            return element.textContent.trim().replace(/\t|\n/g, '');
        }

        let allDataToCopy = ''; // Accumulate data from all rows

        tableRows.forEach(row => {
            const rowDataArray = targetColumnsIndices.map(index => {
                const cellData = row.querySelectorAll('td')[index - 1];
                if (cellData) {
                    let cellText = extractTextContent(cellData);

                    // Exclude "(RMA)" from the end of rowDataArray[3]
                    if (index === 3 && cellText.endsWith('(RMA)')) {
                        cellText = cellText.slice(0, -5); // Remove the last 5 characters
                    }

                    return cellText;
                } else {
                    return '';
                }
            });

            // Combine the data from specific columns into the desired format for each row
            const rowText = `\nSITE:${rowDataArray[7]}\nSN:${rowDataArray[2]}\nMPN:${rowDataArray[10]}\n\https://mobility.amazon.com/part/search?search_string=${rowDataArray[2]}&query=GO&search_type=all&max_rows=50\n`;

            allDataToCopy += rowText; // Accumulate data from all rows
        });

        // Get today's date in MM/DD/YY format
        const today = new Date().toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: '2-digit' }).replace(/\//g, '/');

        // Get the text content from the element with ID 'hello-user' and remove "Hello, "
        const aliasText = document.getElementById('hello-user').textContent.trim().replace('Hello, ', '');

        // Combine all data from rows into the desired format
        const dataToCopy = `Hi Team <@present|Present Members>,\n\nRequesting you to please carry out Sanitization on the below part(s), thank you:\n\n${allDataToCopy}`;

        // Create a temporary textarea to copy the data to the clipboard
        const textArea = document.createElement('textarea');
        textArea.value = dataToCopy;
        textArea.style.position = 'fixed';
        textArea.style.top = '0';
        textArea.style.left = '0';
        textArea.style.width = '2em';
        textArea.style.height = '2em';
        textArea.style.padding = '0';
        textArea.style.border = 'none';
        textArea.style.outline = 'none';
        textArea.style.boxShadow = 'none';
        textArea.style.background = 'transparent';

        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();

        try {
            const successful = document.execCommand('copy');
            if (successful) {
                // Change the button text to indicate successful copying
                copyToClipboardButton.textContent = '💬 ✔️';
                // Show success banner
                CopyMessageSuccessBanner('💬 <strong>Request San message</strong> copied to clipboard - Paste into <strong>Chime / Slack</strong>');
            }
        } catch (err) {
            console.error('Unable to copy to clipboard:', err);
            alert('Failed to copy data to clipboard');
        }

        // Reset the button text after a short delay (e.g., 2 seconds)
        setTimeout(() => {
            copyToClipboardButton.textContent = 'MESSAGE 📋';
        }, 2000); // Adjust the time in milliseconds as needed
    }

    // Function to show the yellow banner with moving truck emoji for 3 seconds
    const CopyMessageSuccessBanner = (CopyMessage) => {
        console.log('Show success banner:', CopyMessage);

        // Create the COPY MSG Success banner element
        const CopyBanner = document.createElement('div');
        CopyBanner.innerHTML = CopyMessage; // Use innerHTML instead of textContent
        CopyBanner.style.position = 'fixed';
        CopyBanner.style.top = '0';
        CopyBanner.style.left = '0';
        CopyBanner.style.width = '100%';
        CopyBanner.style.fontSize = '20px';
        CopyBanner.style.backgroundColor = ' #f1c40f ';
        CopyBanner.style.color = 'black';
        CopyBanner.style.textAlign = 'center';
        CopyBanner.style.padding = '10px';
        CopyBanner.style.zIndex = '9999';
        CopyBanner.style.border = '1px solid black'; // Add black border

        // Append the banner to the document body
        document.body.appendChild(CopyBanner);

        // After 3 seconds, remove the banner
        setTimeout(() => {
            document.body.removeChild(CopyBanner);
        }, 4000);
    };

    // Create the button to trigger the function
    const copyToClipboardButton = document.createElement('button');
    copyToClipboardButton.textContent = 'MESSAGE 📋';
    copyToClipboardButton.onclick = copyDataToClipboard;
    copyToClipboardButton.classList.add('exampleButton');
    copyToClipboardButton.setAttribute('title', 'Creates request message for DCO to perform san on listed SNs'); // Adding tooltip message

    // Append the button to the document body
    document.body.appendChild(copyToClipboardButton);

















    ///////////🔗🔢 COPY SERIALS BUTTON////////////////////////////  *12s to 3s



    function copySerials() {
        const serials = [];
        const tableRows = document.querySelectorAll('table.table-bordered.table-striped.table-condensed tbody tr');

        const extractTextContent = (element) => {
            return element.textContent.trim().replace(/\t|\n/g, '');
        };

        tableRows.forEach(row => {
            const serialCellData = row.querySelectorAll('td')[3 - 1];
            if (serialCellData) {
                let serialText = extractTextContent(serialCellData);

                // Exclude "(RMA)" from the end of serial numbers
                if (serialText.endsWith('(RMA)')) {
                    serialText = serialText.slice(0, -5).trim(); // Remove the last 5 characters
                }

                serials.push(serialText);
            }
        });

        // Copy the serials to the clipboard
        navigator.clipboard.writeText(serials.join('\n'))
            .then(() => {
            console.log('Serials copied to clipboard:', serials.join('\n'));
            const copySerialsButton = document.getElementById('copySerialsButton');
            copySerialsButton.textContent = 'SNs ✔️';
            setTimeout(() => {
                copySerialsButton.textContent = 'SNs 📋';
            }, 500); // Revert back to original text after 2 seconds
        })
            .catch(err => {
            console.error('Unable to copy serials to clipboard:', err);
            alert('Failed to copy serials to clipboard');
        });
    }



    const copySerialsButton = document.createElement('button');
    copySerialsButton.textContent = 'SNs 📋';
    copySerialsButton.id = 'copySerialsButton';
    copySerialsButton.onclick = copySerials;
    copySerialsButton.classList.add('exampleButton');
    copySerialsButton.setAttribute('title', '⚠️ Copies ALL displayed SNs from the table in column format for easy pasting');

















    ////////▶️ 🖥 TRANSFER SERIALS MESSAGE DROPDOWN/////////////



    // Function to transfer serials
    function transferSerials2(option) {
        const serials = [];
        const tableRows = document.querySelectorAll('table.table-bordered.table-striped.table-condensed tbody tr');

        // Function to extract text content from an element
        const extractTextContent = (element) => {
            return element.textContent.trim().replace(/\t|\n/g, '');
        };

        // Define weights for different part types
        const partWeights = {
            'ATS': 12, // ATS type with weight 12 kg
            'Memory': 0.03, // DIMMs (Samsung, etc.) with weight 0.03 kg per module
            'Power Supply': 2, // Power Supply Unit with weight 2 kg
            'Motherboard': 1, // Motherboard with weight 1 kg
            'Processor': 0.1, // CPU (Central Processing Unit) with weight 0.1 kg
            'Fan': 0.1, // Fan with weight 0.1 kg
            'Cable': 0.02, // Cable with weight 0.02 kg per meter
            'Switch': 12, // Switch (Router) with weight 3.5 kg
            'NIC': 0.1, // NIC (Network Interface Card) with weight 0.1 kg
            'Dongle': 0.2, // Dongle with weight 0.03 kg
            'FiberOpticModule': 0.075, // Fiber optic module with weight 0.075 kg
            // Add more part types and their weights here
        };

        // Group serials by MPN (indice 12 from table)
        const groupedSerials = {};
        let totalQty = 0;
        let totalWeight = 0; // Initialize total weight
        tableRows.forEach(row => {
            const mpnCellData = row.querySelectorAll('td')[12 - 1]; // Assuming the MPN is in the twelfth column
            const typeCellData = row.querySelectorAll('td')[17 - 1]; // Assuming the Type is in the seventeenth column
            if (mpnCellData && typeCellData) {
                const mpn = extractTextContent(mpnCellData);
                const type = extractTextContent(typeCellData);
                const serialCellData = row.querySelectorAll('td')[3 - 1]; // Assuming the serials are in the third column
                let serialText = extractTextContent(serialCellData);

                // Exclude "(RMA)" from the end of serial numbers
                if (serialText.endsWith('(RMA)')) {
                    serialText = serialText.slice(0, -5).trim(); // Remove the last 5 characters
                }

                if (!groupedSerials[mpn]) {
                    groupedSerials[mpn] = [];
                }
                groupedSerials[mpn].push({ type, serial: serialText });
                totalQty += 1; // Increment total quantity

                // Add weight of current part type to total weight
                if (partWeights[type]) {
                    totalWeight += partWeights[type];
                }
            }
        });

        // Define the total quantity and total weight
        const QTY = totalQty;
        const WEIGHT = totalWeight.toFixed(1); // Round the weight to one decimal place
        const DIMENSIONS = 'AddHere';


        function downloadCSV(headers, arrayData, fileName, totalQty) {
            // Prepend total quantity to the data
            let csvContent = `Total Qty:,${totalQty}\n`; // Add total quantity as the first row
            csvContent += headers.join(",") + "\n" + arrayData.map(e => e.join(",")).join("\n");
            csvContent = "data:text/csv;charset=utf-8," + csvContent;

            // Create a link and trigger download
            const encodedUri = encodeURI(csvContent);
            const link = document.createElement("a");
            link.setAttribute("href", encodedUri);
            link.setAttribute("download", fileName);
            link.click();
        }





        // Construct the message with ASCII art table layout
        let message = '';

        switch(option) {
            case 'GENERAL 📋':
                var allItems = Object.entries(groupedSerials).flatMap(([mpn, items]) =>
                                                                      items.map(item => ({
                    serial: item.serial,
                    mpn: mpn
                }))
                                                                     );



                var baseURL = 'https://mobility.amazon.com/part/search';
                var maxRows = '6000';
                var query = 'GO';
                var allSerials = allItems.map(item => item.serial).join('+');
                var allSerialsURL = `[${QTY}](${baseURL}?search_type=all&search_string=${encodeURIComponent(allSerials).replace(/%2B/g, '+')}&max_rows=${maxRows}&query=${query}&total_serials=${allSerials.length})`;

                message += 'Transferring below parts:\n-\n---\n';
                message += `Total Qty: **${allSerialsURL}**\n`;
                message += `Total Weight: **${WEIGHT}** kg\n`;
                message += `Dimensions: **${DIMENSIONS}**\n`;
                message += `Tracking ID: **AddHere**\n`;
                message += '________________________________________________________________________________________________________________________________\n\n';

                // Existing loop for appending message content
                for (const mpn in groupedSerials) {
                    message += '>\n';
                    message += 'TYPE: **' + groupedSerials[mpn][0].type + '**\n';
                    message += 'MPN: **' + mpn + '**\n';
                    message += 'QTY: **' + groupedSerials[mpn].length + '**\n\n';

                    const queryString = groupedSerials[mpn].map(item => item.serial).filter(Boolean).join('+');
                    const url = `**[Mobility 🔗](${baseURL}?search_type=all&search_string=${encodeURIComponent(queryString).replace(/%2B/g, '+')}&max_rows=${maxRows}&query=${query})**`;

                    message += url;
                    message += '\n\n';
                    message += '**SERIAL #:**\n\n' + groupedSerials[mpn].map(item => '    ' + item.serial + '    ').join('\n') + '\n\n';
                    message += '________________________________________________________________________________________________________________________________\n\n';
                }
                break;



            case 'RZPR ⛔':
                // Group serials by Type and MPN
                var groupedSerialsByType = {};

                // Get Operator value
                var operatorElement = document.getElementById('hello-user');
                var operator = operatorElement.textContent.split(',')[1].trim(); // Extract the text after "Hello, "

                tableRows.forEach(row => {
                    const typeCellData = row.querySelectorAll('td')[17 - 1]; // Assuming the Type is in the seventeenth column
                    const serialCellData = row.querySelectorAll('td')[3 - 1]; // Assuming the serials are in the third column
                    const mpnCellData = row.querySelectorAll('td')[12 - 1]; // Assuming the MPN is in the twelfth column

                    if (typeCellData && serialCellData && mpnCellData) {
                        const type = extractTextContent(typeCellData);
                        let serial = extractTextContent(serialCellData);
                        let mpn = extractTextContent(mpnCellData);

                        // Exclude "(RMA)" from the end of serial numbers
                        if (serial.endsWith('(RMA)')) {
                            serial = serial.slice(0, -5).trim();
                        }

                        if (!groupedSerialsByType[type]) {
                            groupedSerialsByType[type] = [];
                        }
                        groupedSerialsByType[type].push({ serial, mpn });
                    }
                });

                // Create a sorted array of types
                var sortedTypes = Object.keys(groupedSerialsByType).sort();

                // Flatten all serials into a single array
                allSerials = sortedTypes.flatMap(type => groupedSerialsByType[type].map(item => item.serial)).join('+');

                // Generate the URL for all serials
                baseURL = 'https://mobility.amazon.com/part/search';
                maxRows = '6000';
                query = 'GO';
                allSerialsURL = `[${Object.values(groupedSerialsByType).reduce((acc, val) => acc + val.length, 0)}](${baseURL}?search_type=all&search_string=${encodeURIComponent(allSerials).replace(/%2B/g, '+')}&max_rows=${maxRows}&query=${query}&total_serials=${allSerials.length})`;

                // Construct the summary box
                message += 'Removing Below Serials ⛔:\n-\n';
                message += `**Approval ✅**: AddLinkHere\n\n---\n`;
                message += `**Seals 🔒**: \nAddSealsHere\n\n---\n`;
                message += `    Operator: ${operator}\n`;
                message += `    Controller: AddAliasHere\n`;
                message += `    Trustee: AddAliasHere\n\n`;
                message += `Total Qty: **${allSerialsURL}**\n-\n`;
                sortedTypes.forEach(type => {
                    message += `    ${type.toUpperCase()}: ${groupedSerialsByType[type].length}\n`;
                });
                message += '\n';

                // Construct the main table
                message += 'Serial | Type | MPN\n';
                message += '---|---|---\n';

                sortedTypes.forEach(type => {
                    groupedSerialsByType[type].forEach(item => {
                        message += `${item.serial} | ${type.toUpperCase()} | ${item.mpn}\n`;
                    });
                });

                message += '---\n';

                break;







            case 'TOTE 📦': {


                // Define a function expression to group serials by toteId
                const groupSerialsByToteIdFunction = (tableRows) => {
                    const groupedSerialsByToteId = {};
                    tableRows.forEach(row => {
                        const toteIdCellData = row.querySelectorAll('td')[24 - 1]; // Assuming toteId is in the twenty-fourth column
                        const serialCellData = row.querySelectorAll('td')[3 - 1]; // Assuming the serials are in the third column
                        if (toteIdCellData && serialCellData) {
                            const toteId = extractTextContent(toteIdCellData);
                            const serialText = extractTextContent(serialCellData);

                            // Exclude "(RMA)" from the end of serial numbers
                            let serial = serialText.endsWith('(RMA)') ? serialText.slice(0, -5).trim() : serialText;

                            if (!groupedSerialsByToteId[toteId]) {
                                groupedSerialsByToteId[toteId] = [];
                            }
                            groupedSerialsByToteId[toteId].push(serial);
                        }
                    });
                    return groupedSerialsByToteId;
                };

                // Group serials by toteId
                const groupedSerialsByToteId = groupSerialsByToteIdFunction(tableRows);

                // Start building the message
                message += 'Part(s) ready for collection from\n---\n';
                var uniqueBinIds = new Set(); // To ensure unique binIds
                tableRows.forEach(row => {
                    const originSiteCellData = row.querySelectorAll('td')[8 - 1]; // Assuming OriginSite is in the eighth column
                    const toteIdCellData = row.querySelectorAll('td')[24 - 1]; // Assuming toteId is in the twenty-fourth column
                    const binNameCellData1 = row.querySelectorAll('td')[8 - 1]; // Assuming binName1 is in the eighth column
                    const binNameCellData2 = row.querySelectorAll('td')[9 - 1]; // Assuming binName2 is in the ninth column
                    const binNameCellData3 = row.querySelectorAll('td')[10 - 1]; // Assuming binName3 is in the tenth column

                    if (originSiteCellData && toteIdCellData && binNameCellData1 && binNameCellData2 && binNameCellData3) {
                        const originSite = extractTextContent(originSiteCellData);
                        const toteId = extractTextContent(toteIdCellData);
                        const binName = extractTextContent(binNameCellData1) + '.' + extractTextContent(binNameCellData2) + '.' + extractTextContent(binNameCellData3);

                        // Append to message if binId is unique
                        if (!uniqueBinIds.has(binName)) {
                            message += `**${originSite}** to **SITE**\n\n`;
                            message += `📦 Tote Name: **[${toteId}](https://mobility.amazon.com/part/search?search_type=all&search_string=${toteId}&max_rows=500&query=GO)**\n\n`;
                            message += `📍 Location: **${binName}**\n\n\n`;
                            message += '\n\n';

                            // Get all serials associated with this toteId
                            const toteSerials = (groupedSerialsByToteId[toteId] || []).map(serial => `    ${serial}`).join('\n');
                            message += `**SERIAL IDs**:\n\n${toteSerials || 'No serials associated with this tote.'}\n`;

                            // Generate URL for QTY
                            const qtyURL = `[QTY: x${groupedSerialsByToteId[toteId].length}](https://mobility.amazon.com/part/search?search_type=all&search_string=${toteId}&max_rows=500&query=GO)`;
                            message += `>\n${qtyURL}\n\n---\n`;

                            uniqueBinIds.add(binName);
                        }
                    }
                });

            }
                break;





                ////FA LAB OPTION //////




            case 'FA LAB 🔎':
                // Define serialsWithoutContainerId to store serials without container ID
                var serialsWithoutContainerId = [];

                // Group serials by type for Option 2
                groupedSerialsByType = {};
                tableRows.forEach(row => {
                    const typeCellData = row.querySelectorAll('td')[17 - 1]; // Assuming the Type is in the seventeenth column
                    if (typeCellData) {
                        const type = extractTextContent(typeCellData);
                        const serialCellData = row.querySelectorAll('td')[3 - 1]; // Assuming the serials are in the third column
                        let serialText = extractTextContent(serialCellData);

                        // Exclude "(RMA)" from the end of serial numbers
                        if (serialText.endsWith('(RMA)')) {
                            serialText = serialText.slice(0, -5).trim(); // Remove the last 5 characters
                        }

                        // Check if the serial has a corresponding bulk container ID
                        const bulkContainerIdCell = row.querySelectorAll('td')[31 - 1]; // Assuming the bulk container ID is in the thirty-first column
                        if (!bulkContainerIdCell) {
                            serialsWithoutContainerId.push(serialText); // Add serial to the list if no bulk container ID is found
                        }

                        if (!groupedSerialsByType[type]) {
                            groupedSerialsByType[type] = [];
                        }
                        groupedSerialsByType[type].push(serialText);
                    }
                });


                // Construct the bulk container IDs section
                var bulkContainerIds = new Set(); // Using a Set to avoid duplicate values
                tableRows.forEach(row => {
                    const bulkContainerIdCell = row.querySelectorAll('td')[31 - 1]; // Assuming the bulk container ID is in the thirty-first column
                    if (bulkContainerIdCell) {
                        const bulkContainerId = extractTextContent(bulkContainerIdCell);
                        bulkContainerIds.add(bulkContainerId);
                    }
                });

                var bulkContainerIdURLs = []; // Store URLs for each bulkContainerId
                bulkContainerIds.forEach(bulkContainerId => {
                    const shavedBulkContainerId = bulkContainerId.substring(5); // Shave off the first 3 letters
                    const bulkContainerIdURL = `https://t.corp.amazon.com/${shavedBulkContainerId}/communication`;
                    bulkContainerIdURLs.push(`[${bulkContainerId}](https://t.corp.amazon.com/${shavedBulkContainerId}/communication)`);
                });

                var bulkContainerIdMessage = bulkContainerIdURLs.join(' | '); // Join URLs with ' | '


                message += '\n\n';

                // Define the base URL
                var baseUrl = 'https://mobility.amazon.com/part/search?search_type=all&max_rows=6000&query=GO';

                // Add bulkContainerId URLs to the message
                message += '**Bulk Container ID(s)**: ';
                bulkContainerIdURLs.forEach((url, index) => {
                    const bulkContainerId = Array.from(bulkContainerIds)[index]; // Get the full bulk container ID
                    const customUrl = `${baseUrl}&search_string=${encodeURIComponent(bulkContainerId)}`;
                    message += `[**${bulkContainerId}**](${customUrl}) | `;
                });
                message += '\n-\n';




                // Get Operator value
                operatorElement = document.getElementById('hello-user');
                operator = operatorElement.textContent.split(',')[1].trim(); // Extract the text after "Hello, "

                // Add Operator and Controller fields
                message += `**Seals 🔒**: \nAddSealsHere\n\n---\n`;
                message += `    Operator: ${operator}\n`;
                message += `    Controller: AddAliasHere\n\n`;

                // Serials ID section
                message += 'Part & Serial ID Summary:\n---\n';

                // Extract type names from groupedSerialsByType keys
                var typeNames = Object.keys(groupedSerialsByType);

                // Construct type names row
                var typeNamesRow = '|';
                for (const type of typeNames) {
                    const uppercaseType = type.toUpperCase(); // Convert type to uppercase
                    typeNamesRow += ` ${uppercaseType} (${groupedSerialsByType[type].length}) |`;
                }
                message += `${typeNamesRow}\n`;

                // Construct row separator
                var rowSeparator = '|';
                for (let i = 0; i < typeNames.length; i++) {
                    rowSeparator += '-|';
                }
                message += `${rowSeparator}\n`;

                var maxSerialsCount = Math.max(...typeNames.map(type => groupedSerialsByType[type].length));

                // Construct rows with serials for each type
                for (let i = 0; i < maxSerialsCount; i++) {
                    let row = '|';
                    for (const type of typeNames) {
                        const serial = groupedSerialsByType[type][i] || ''; // Use empty string if no serial exists at this index
                        row += ` ${serial} |`;
                    }
                    message += `${row}\n`;
                }

                // Add total quantity of serials
                message += `${rowSeparator}\n`;
                var allSerialsCount = Object.values(groupedSerialsByType).reduce((acc, val) => acc + val.length, 0);
                message += `TOTAL QTY 📊: ${allSerialsCount}\n`;
                message += `- \n`;

                // Add final pieces to the message
                message += '\n*All serials match on [Mobility]()*\n\n';
                var types = Object.keys(groupedSerialsByType).map(type => type.toUpperCase()).join(', ');
                message += `All devices in this container have been identified as [**${types}**] and are packed/sealed for transfer to WW for Failure Analysis via the [Approved SOP🔗](https://w.amazon.com/bin/view/WinstonWolfe/FAPartsPickup).\n`;

                break;




        }

        // Define bulkContainerIds and bulkContainerIdURLs outside the switch statement
        bulkContainerIds = new Set();
        bulkContainerIdURLs = [];

        // Construct the bulk container IDs section and populate bulkContainerIdURLs
        tableRows.forEach(row => {
            const bulkContainerIdCell = row.querySelectorAll('td')[31 - 1]; // Assuming the bulk container ID is in the thirty-first column
            if (bulkContainerIdCell) {
                const bulkContainerId = extractTextContent(bulkContainerIdCell);
                const shavedBulkContainerId = bulkContainerId.substring(5); // Shave off the first 3 letters
                const bulkContainerIdURL = `https://t.corp.amazon.com/${shavedBulkContainerId}/communication`;
                bulkContainerIdURLs.push(bulkContainerIdURL); // Add URL for each bulkContainerId
                bulkContainerIds.add(bulkContainerId);
            }
        });

        // Calculate the shaved bulkContainerId
        const shavedBulkContainerIds = Array.from(bulkContainerIds).map(id => id.substring(5));

        // Construct the image element
        var image = document.createElement('img');
        image.src = 'https://drive-render.corp.amazon.com/view/pjbyrne@/Script%20Banner%20Files/Markdown_On.JPG'; // Replace 'image_url_here' with the URL of the image








        //📋 COPY MESSAGE CODE
        navigator.clipboard.writeText(message)
            .then(() => {
            console.log('Message copied to clipboard:', message);
            // Show success banner with custom message and image
            switch(option) {
                case 'GENERAL 📋':
                    showCopyMessageSuccessBanner(`<b>General Transfer 📋</b> Message Ready! ✔️ *Paste into TT and select ${image.outerHTML}`);
                    break;
                case 'TOTE 📦':
                    // Show success banner with custom message and image
                    showCopyMessageSuccessBanner(`<b>Tote Collection 📦</b> Message Ready! ✔️ *Paste into TT and select ${image.outerHTML}`);
                    break;
                case 'FA LAB 🔎':
                    // Collect serials without a containerId
                    var serialsWithoutContainerId = [];
                    tableRows.forEach(row => {
                        const bulkContainerIdCell = row.querySelectorAll('td')[31 - 1]; // Assuming the bulk container ID is in the thirty-first column
                        const serialCellData = row.querySelectorAll('td')[3 - 1]; // Assuming the serials are in the third column
                        if (serialCellData) {
                            let serialText = extractTextContent(serialCellData);
                            // Exclude "(RMA)" from the end of serial numbers
                            if (serialText.endsWith('(RMA)')) {
                                serialText = serialText.slice(0, -5).trim(); // Remove the last 5 characters
                            }
                            if (!bulkContainerIdCell || !bulkContainerIdCell.textContent.trim()) {
                                serialsWithoutContainerId.push(serialText);
                            }
                        }
                    });

                    // Check if there are serials without a container ID
                    if (serialsWithoutContainerId.length > 0) {
                        showErrorMessageBanner('⚠️ Following Serials don\'t have ContainerId:', serialsWithoutContainerId);
                    } else if (bulkContainerIds.size > 1) {
                        // Check if more than 1 bulkContainerId is detected
                        showErrorMessageBanner('⚠️ More than 1 Container ID detected!', bulkContainerIds);
                    } else {
                        // Include hyperlinks for bulkContainerIds after "Ticket!"
                        var bulkContainerIdMessage = bulkContainerIdURLs.map((url, index) => `[Ticket${index + 1}: ${url}]`).join(' | ');

                        // Construct a single hyperlink for all serials
                        const allSerialsUrl = `https://mobility.amazon.com/part/search?search_type=all&search_string=${encodeURIComponent(groupedSerialsByType[typeNames[0]].join('+'))}&max_rows=6000&query=GO`;
                        const allSerialsLink = `<a href="${allSerialsUrl}" target="_blank">${allSerialsCount} Serials</a>`;

                        // Append allSerialsLink only when both conditions are met
                        showCopyMessageSuccessBanner(`<b>FA Lab Transfer 🔎</b> Message Ready! ✔️ *Please paste into <span style="color: blue;">${shavedBulkContainerIds.map((id, index) => `<a href="https://t.corp.amazon.com/${id}/communication" target="_blank" style="color: blue;">${id}</a>`).join(' | ')}</span>, add <b>Seals</b>🔒 & <b>Controller</b> and select ${image.outerHTML}\n\n`);
                    }

                    break;

                case 'RZPR ⛔':
                    showCopyMessageSuccessBanner(`<b>RZPR ⛔ Transfer</b> Message Ready! ✔️`);
                    break;

                default:
                    showCopyMessageSuccessBanner('Transfer successful!');
            }
        })
            .catch(err => {
            console.error('Unable to copy message to clipboard:', err);
            alert('Failed to copy message to clipboard');
        });

    }





    ///////  ✅ SUCCESS BANNER
    const showCopyMessageSuccessBanner = (message) => {
        const successBanner = document.createElement('div');
        successBanner.innerHTML = message; // Use innerHTML instead of textContent
        successBanner.style.position = 'fixed';
        successBanner.style.top = '0';
        successBanner.style.left = '0';
        successBanner.style.width = '100%';
        successBanner.style.fontSize = '20px';
        successBanner.style.fontFamily = 'roboto';
        successBanner.style.backgroundColor = '#e5e7e9'; // Green color for success
        successBanner.style.color = 'black'; // White text color
        successBanner.style.textAlign = 'center';
        successBanner.style.padding = '10px';
        successBanner.style.zIndex = '9999';
        successBanner.style.border = '1px solid black'; // Add black border

        // Append the banner to the document body
        document.body.appendChild(successBanner);

        // After 2 seconds, remove the banner
        setTimeout(() => {
            document.body.removeChild(successBanner);
        }, 8000);
    };

    //////  ⛔ ERROR BANNER

    const showErrorMessageBanner = (errorMessage, bulkContainerIds) => {
        const errorBanner = document.createElement('div');
        errorBanner.innerHTML = errorMessage; // Use innerHTML instead of textContent
        errorBanner.style.position = 'fixed';
        errorBanner.style.top = '0';
        errorBanner.style.left = '0';
        errorBanner.style.width = '100%';
        errorBanner.style.fontSize = '20px';
        errorBanner.style.fontFamily = 'roboto';
        errorBanner.style.backgroundColor = '#e74c3c'; // Red color for error
        errorBanner.style.color = 'black'; // White text color
        errorBanner.style.textAlign = 'center';
        errorBanner.style.padding = '10px';
        errorBanner.style.zIndex = '9999';
        errorBanner.style.border = '1px solid black'; // Add black border

        // Create a list of hyperlinks for bulkContainerIds
        const bulkContainerLinks = Array.from(bulkContainerIds).map(bulkContainerId => {
            const url = `https://mobility.amazon.com/part/search?search_type=all&search_string=${encodeURIComponent(bulkContainerId)}&max_rows=6000&query=GO`;
            return `<a href="${url}" target="_blank">${bulkContainerId}</a>`;
        }).join(' | ');

        // Append the list to the error banner
        const containerIdsDiv = document.createElement('div');
        containerIdsDiv.innerHTML = bulkContainerLinks;
        errorBanner.appendChild(containerIdsDiv);

        // Append the banner to the document body
        document.body.appendChild(errorBanner);

        // After 8 seconds, remove the banner
        setTimeout(() => {
            document.body.removeChild(errorBanner);
        }, 8000);
    };





    // Create the dropdown for transfer options
    const transferOptionsDropdown = document.createElement('select');
    transferOptionsDropdown.id = 'transferOptionsDropdown';

    // Apply styles to the dropdown
    transferOptionsDropdown.style.backgroundColor = '#31465E'; // Set background color
    transferOptionsDropdown.style.border = '1px solid black'; // Add black border
    transferOptionsDropdown.style.height = '36px'; // Set height
    transferOptionsDropdown.style.width = '130px'; // Set width
    transferOptionsDropdown.style.color = 'white'; // Set font color
    transferOptionsDropdown.style.padding = '6px 16px'; // Adjust padding
    transferOptionsDropdown.style.textAlign = 'center'; // Center text horizontally
    transferOptionsDropdown.style.textDecoration = 'none'; // Remove underline
    transferOptionsDropdown.style.fontSize = '14px'; // Set font size
    transferOptionsDropdown.style.fontWeight = 'bold'; // Set font weight to bold
    transferOptionsDropdown.style.margin = '2px'; // Adjust margin
    transferOptionsDropdown.style.cursor = 'pointer'; // Set cursor to pointer
    transferOptionsDropdown.style.borderRadius = '5px'; // Set border radius
    transferOptionsDropdown.style.transition = 'background-color 0.3s, box-shadow 0.3s'; // Add transition effect for background color and box-shadow
    transferOptionsDropdown.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Add box shadow for 3D effect
    transferOptionsDropdown.setAttribute('title', '⚠️ Various options to copy displayed serials in markdown format for ticket correspondence'); // Adding tooltip message

    // Add hover effect
    transferOptionsDropdown.addEventListener('mouseenter', function () {
        transferOptionsDropdown.style.backgroundColor = ' #5d6d7e ';
        transferOptionsDropdown.style.color = 'white';
        transferOptionsDropdown.style.boxShadow = '0px 8px 10px rgba(0, 0, 0, 0.4)'; // Adjust shadow properties for the hover effect
    });

    transferOptionsDropdown.addEventListener('mouseleave', function () {
        transferOptionsDropdown.style.backgroundColor = '#31465E';
        transferOptionsDropdown.style.color = 'white';
        transferOptionsDropdown.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Reset shadow properties
    });

    // Add default label "Select Transfer"
    const defaultOption1 = document.createElement('option');
    defaultOption1.text = 'TRANSFER💬';
    defaultOption1.disabled = true;
    defaultOption1.selected = true;
    defaultOption1.style.fontFamily = 'Roboto'; // Set font family to Roboto for the default option
    transferOptionsDropdown.appendChild(defaultOption1);

    // Add transfer options
    const options = ['GENERAL 📋', 'RZPR ⛔', 'TOTE 📦', 'FA LAB 🔎'];
    options.forEach(option => {
        const optionElement = document.createElement('option');
        optionElement.value = option;
        optionElement.text = option;
        optionElement.style.fontFamily = 'Roboto'; // Set font family to Roboto for each option
        transferOptionsDropdown.appendChild(optionElement);
    });



    // Event listener for dropdown change
    transferOptionsDropdown.addEventListener('change', (event) => {
        const selectedOption = event.target.value;
        transferSerials2(selectedOption);
        event.target.value = 'TRANSFER💬'; // Reset selected option to default
    });








    //////📏📋 COPY SHORTEN URL //////



    // Function to shorten the URL using TinyURL API
    function shortenUrl() {
        let currentUrl = window.location.href;

        // Define the segments to remove
        const segmentsToRemove = ['/mobiamazpartsear', '/produseamobiscmawsdevpart'];

        // Check and remove the segment if it exists at the end of the URL
        segmentsToRemove.forEach(segment => {
            if (currentUrl.endsWith(segment)) {
                currentUrl = currentUrl.slice(0, -segment.length);
                console.log(`🔍 "${segment}" segment removed from URL:`, currentUrl);
            }
        });

        const apiUrl = 'https://tiny.AMAZON.com/submit/url?name=';

        // Perform a CORS request to the TinyURL API
        GM_xmlhttpRequest({
            method: 'GET',
            url: apiUrl + encodeURIComponent(currentUrl),
            onload: function(response) {
                if (response.status === 200) {
                    // Extract the shortened URL from the response
                    const responseText = response.responseText.trim(); // Remove any leading/trailing whitespace

                    // Check if the response contains the shortened URL
                    const match = responseText.match(/href="([^"]+)"/);
                    if (match && match[1]) {
                        let shortenedUrl = match[1];

                        // Check and remove the segments again in case they are reintroduced
                        segmentsToRemove.forEach(segment => {
                            if (shortenedUrl.includes(segment)) {
                                shortenedUrl = shortenedUrl.replace(segment, '');
                                console.log(`🔍 "${segment}" segment removed from shortened URL:`, shortenedUrl);
                            }
                        });

                        // Copy the cleaned shortened URL to clipboard
                        copyShortURL(shortenedUrl);

                        // Temporarily change the button text to indicate success
                        shortenButton.innerHTML = '<strong>URL✔️</strong>';
                        setTimeout(() => {
                            shortenButton.innerHTML = '<strong>URL📉</strong>';
                        }, 1000); // Revert back to original text after 1000ms
                    } else {
                        console.error('Shortened URL not found in response:', responseText);
                    }
                } else {
                    console.error('Error:', response.statusText);
                }
            },
            onerror: function(error) {
                console.error('Error:', error);
            }
        });
    }

    // Function to copy text to clipboard
    function copyShortURL(text) {
        const tempInput = document.createElement('input');
        tempInput.value = text;
        document.body.appendChild(tempInput);
        tempInput.select();
        document.execCommand('copy');
        document.body.removeChild(tempInput);

        // Display success message
        const successBanner = document.createElement('div');
        successBanner.innerHTML = 'URL📉 Copied 📋: <strong>' + text + '</strong>';
        successBanner.style.color = 'black';
        successBanner.style.position = 'fixed';
        successBanner.style.top = '0';
        successBanner.style.left = '0';
        successBanner.style.width = '100%';
        successBanner.style.fontSize = '20px';
        successBanner.style.backgroundColor = '#e5e7e9';
        successBanner.style.color = 'black';
        successBanner.style.textAlign = 'center';
        successBanner.style.padding = '10px';
        successBanner.style.zIndex = '9999';
        successBanner.style.border = '1px solid black'; // Add black border
        document.body.appendChild(successBanner);

        // Remove the success message after 3 seconds
        setTimeout(function() {
            document.body.removeChild(successBanner);
        }, 5000);
    }

    // Create a button to trigger the URL shortening process
    const shortenButton = document.createElement('button');
    shortenButton.innerHTML = '<strong>URL📉</strong>';
    shortenButton.id = 'shortenButton';
    shortenButton.onclick = shortenUrl;
    shortenButton.classList.add('exampleButton');
    shortenButton.setAttribute('title', '⚠️ Shortens URL and copies it to clipboard');







    /////////////////🧱 ADD BUILD ID BUTTON 🧱/////////////////////////////////////   12s > 6s  /  6 click > 3 click



    // Define the input textbox globally
    const BuildInputTextbox = document.createElement('input');

    function BuildIdInput() {
        BuildInputTextbox.type = 'text';
        BuildInputTextbox.placeholder = 'BUILD 🏗️';
        BuildInputTextbox.style.textAlign = 'center'; // Center text



        // Add hover description
        BuildInputTextbox.title = '⚠️ Sets checked ☑️ SNs BUILD ID';

        // Function to show the yellow banner with moving truck emoji for 3 seconds
        const BuildIdSuccessBanner = (BuildMessage) => {
            console.log('Show success banner:', BuildMessage);

            // Create the BUILD Success banner element
            const BuildBanner = document.createElement('div');
            BuildBanner.innerHTML = BuildMessage; // Use innerHTML instead of textContent
            BuildBanner.style.position = 'fixed';
            BuildBanner.style.top = '0';
            BuildBanner.style.left = '0';
            BuildBanner.style.width = '100%';
            BuildBanner.style.fontSize = '20px';
            BuildBanner.style.backgroundColor = '#e5e7e9';
            BuildBanner.style.color = 'black';
            BuildBanner.style.textAlign = 'center';
            BuildBanner.style.padding = '10px';
            BuildBanner.style.zIndex = '9999';
            BuildBanner.style.border = '1px solid black'; // Add black border


            // Append the banner to the document body
            document.body.appendChild(BuildBanner);

            // After 3 seconds, remove the banner
            setTimeout(() => {
                document.body.removeChild(BuildBanner);
            }, 5000);
        };

        // Function to show the red error banner for 6 seconds
        const BuildErrorBanner = (errorMessage) => {
            console.log('Show error banner:', errorMessage);

            // Create the red banner element
            const BuildIdErrorBanner = document.createElement('div');
            BuildIdErrorBanner.innerHTML = errorMessage;
            BuildIdErrorBanner.style.position = 'fixed';
            BuildIdErrorBanner.style.top = '0';
            BuildIdErrorBanner.style.left = '0';
            BuildIdErrorBanner.style.width = '100%';
            BuildIdErrorBanner.style.fontSize = '20px';
            BuildIdErrorBanner.style.backgroundColor = 'red';
            BuildIdErrorBanner.style.color = 'white';
            BuildIdErrorBanner.style.textAlign = 'center';
            BuildIdErrorBanner.style.padding = '10px';
            BuildIdErrorBanner.style.zIndex = '9999';
            BuildIdErrorBanner.style.border = '1px solid black'; // Add black border

            // Append the banner to the document body
            document.body.appendChild(BuildIdErrorBanner);

            // After 6 seconds, remove the banner
            setTimeout(() => {
                document.body.removeChild(BuildIdErrorBanner);
            }, 3000);
        };



        function changeBuildId() {
            BuildInputTextbox.addEventListener('keydown', function (event) {
                if (event.keyCode === 13) { // Check if Enter key is pressed
                    const BuildId = BuildInputTextbox.value.trim();
                    console.log('Entered build ID:', BuildId);

                    // Check if the input is empty
                    if (BuildId === '') {
                        // Display error message if no input is provided
                        BuildErrorBanner('⚠️ Please enter <strong>Build ID!</strong>');
                        return;
                    }

                    // Trigger the action to select checkboxes (like clicking a button)
                    // For example:
                    // makeChangesButton.click();

                    // After triggering the action to select checkboxes, check if any checkboxes are checked
                    const checkboxElements = document.querySelectorAll('input[type="checkbox"][title="Select checkbox to mark this part for Bulk Edit"]');
                    const checkboxesChecked = Array.from(checkboxElements).filter(checkbox => checkbox.checked);

                    if (checkboxesChecked.length === 0) {
                        // Display error message if no row has been selected
                        console.log('⚠️ Please select parts to change!');
                        BuildErrorBanner('⚠️ Please select☑️ parts to change!');
                        return;
                    }

                    // Log the number of selected rows
                    console.log('Number of rows to change:', checkboxesChecked.length);

                    // Select build ID checkbox
                    const BuildIdCheckbox = document.getElementById('edit_build_id_chkbox');
                    if (BuildIdCheckbox) {
                        BuildIdCheckbox.checked = true;
                        console.log('Build ID checkbox checked');

                        // Append text to the build ID textbox
                        const BuildIdTextbox = document.getElementById('build_id_text');
                        if (BuildIdTextbox) {
                            // Append the entered build ID to the build ID textbox
                            BuildIdTextbox.value = BuildId;
                            console.log('Build ID Textbox value set:', BuildId);

                            // Trigger the "Make Changes" button using XPath
                            const container = document.getElementById('part_search_bulk_edit_popup_dialog_form');
                            if (container) {
                                const makeChangesButtonXPath = `//div[@class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix"]//button[contains(text(), "Make Changes")]`;
                                const makeChangesButton = document.evaluate(makeChangesButtonXPath, container, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

                                if (makeChangesButton) {
                                    makeChangesButton.click();
                                    console.log('Make Changes button clicked');

                                    // Function to periodically check the background color of selected checkbox elements
                                    const checkBackgroundColorBuild = setInterval(() => {
                                        let redCheckboxes = 0;
                                        let greenCheckboxes = 0;
                                        checkboxElements.forEach((checkboxElement) => {
                                            // Get the parent td element
                                            const tdElement = checkboxElement.closest('td');
                                            // Log the background color of the td element
                                            if (tdElement && window.getComputedStyle(tdElement).getPropertyValue('background-color') === 'rgb(50, 205, 50)') {
                                                greenCheckboxes++;
                                            } else if (tdElement && window.getComputedStyle(tdElement).getPropertyValue('background-color') === 'rgb(255, 0, 0)') {
                                                redCheckboxes++;
                                            }
                                        });

                                        if (greenCheckboxes === checkboxesChecked.length) {
                                            // All checkboxes have turned green, display success message
                                            BuildIdSuccessBanner(`🏗️ x<strong>${checkboxesChecked.length}</strong> Part(s) <strong>Build ID</strong> updated to: <strong>${BuildId}</strong>`);
                                            clearInterval(checkBackgroundColorBuild);

                                            // Reset cell colors
                                            checkboxElements.forEach((checkboxElement) => {
                                                const tdElement = checkboxElement.closest('td');
                                                if (tdElement) {
                                                    tdElement.style.backgroundColor = 'inherit';
                                                }
                                            });

                                            // Uncheck all selected checkboxes
                                            checkboxesChecked.forEach((checkboxElement) => {
                                                checkboxElement.checked = false;
                                            });

                                            // Reset everything in the form
                                            BuildIdCheckbox.checked = false;
                                            BuildIdTextbox.value = '';

                                            // Clear the input textbox
                                            BuildInputTextbox.value = '';
                                            // Uncheck the "Bulk Edit Select All" checkbox if it was selected before running the function
                                            const bulkEditSelectAllCheckbox = document.getElementById('part_search_bulk_edit_select_all');
                                            if (bulkEditSelectAllCheckbox && bulkEditSelectAllCheckbox.checked) {
                                                bulkEditSelectAllCheckbox.checked = false;
                                                console.log('Bulk edit select all checkbox unchecked');
                                            }
                                        } else if (redCheckboxes > 0 && redCheckboxes + greenCheckboxes === checkboxesChecked.length) {
                                            // Display error message if any cell is red and all cells have changed color
                                            BuildErrorBanner(`⚠️ Error: Please reload page and try again`);
                                            clearInterval(checkBackgroundColorBuild);

                                            // Reset cell colors
                                            checkboxElements.forEach((checkboxElement) => {
                                                const tdElement = checkboxElement.closest('td');
                                                if (tdElement) {
                                                    tdElement.style.backgroundColor = 'inherit';
                                                }
                                            });

                                            // Uncheck all selected checkboxes
                                            checkboxesChecked.forEach((checkboxElement) => {
                                                checkboxElement.checked = false;
                                            });

                                            // Reset everything in the form
                                            BuildIdCheckbox.checked = false;
                                            BuildIdTextbox.value = '';
                                            // Clear the input textbox
                                            BuildInputTextbox.value = '';


                                        }
                                    }, 500); // Check every 0.5 seconds
                                } else {
                                    console.log('Make Changes button not found!');
                                }
                            } else {
                                console.log('Container element not found!');
                            }
                        } else {
                            console.log('Build ID Textbox not found!');
                        }
                    } else {
                        console.log('Build ID checkbox not found!');
                    }
                }
            });
        }


        // Call the function to listen for changes and "Enter" key press
        changeBuildId();


    }

    // Call the function to create the input textbox and listen for changes and "Enter" key press
    BuildIdInput();

    // Append the input textbox to the document body
    document.body.appendChild(BuildInputTextbox);










    /////////////////↩️ RMA ID BUTTON/////////////////////// *Time/Step reduction: 12s to 6s



    // Define the input textbox globally
    const RmaIdInputTextbox = document.createElement('input');

    function RmaIdInput() {
        RmaIdInputTextbox.type = 'text';
        RmaIdInputTextbox.placeholder = 'RMA ID ↩️';
        RmaIdInputTextbox.style.textAlign = 'center'; // Center text



        // Add hover description
        RmaIdInputTextbox.title = '⚠️ Sets checked ☑️ SNs to RMA_PROCESSED';



        // Function to show the blue banner for RMA success
        const RmaSuccessBanner = (RmaSuccessMessage) => {
            console.log('Show success banner:', RmaSuccessMessage);

            // Create the RMA Success banner element
            const RmaBanner = document.createElement('div');
            RmaBanner.innerHTML = RmaSuccessMessage;
            RmaBanner.style.position = 'fixed';
            RmaBanner.style.top = '0';
            RmaBanner.style.left = '0';
            RmaBanner.style.width = '100%';
            RmaBanner.style.fontSize = '20px';
            RmaBanner.style.backgroundColor = '#5dade2';
            RmaBanner.style.color = 'white';
            RmaBanner.style.textAlign = 'center';
            RmaBanner.style.padding = '10px';
            RmaBanner.style.zIndex = '9999';
            RmaBanner.style.border = '1px solid black'; // Add black border

            // Append the banner to the document body
            document.body.appendChild(RmaBanner);

            // After 3 seconds, remove the banner
            setTimeout(() => {
                document.body.removeChild(RmaBanner);
            }, 3000);
        };

        // Function to show the red error banner for 6 seconds
        const RmaErrorBanner = (errorMessage) => {
            console.log('Show error banner:', errorMessage);

            // Create the red banner element
            const RmaErrorBanner = document.createElement('div');
            RmaErrorBanner.innerHTML = errorMessage;
            RmaErrorBanner.style.position = 'fixed';
            RmaErrorBanner.style.top = '0';
            RmaErrorBanner.style.left = '0';
            RmaErrorBanner.style.width = '100%';
            RmaErrorBanner.style.fontSize = '20px';
            RmaErrorBanner.style.backgroundColor = 'red';
            RmaErrorBanner.style.color = 'white';
            RmaErrorBanner.style.textAlign = 'center';
            RmaErrorBanner.style.padding = '10px';
            RmaErrorBanner.style.zIndex = '9999';
            RmaErrorBanner.style.border = '1px solid black'; // Add black border

            // Append the banner to the document body
            document.body.appendChild(RmaErrorBanner);

            // After 6 seconds, remove the banner
            setTimeout(() => {
                document.body.removeChild(RmaErrorBanner);
            }, 3000);
        };

        function changeOutboundRmaId() {
            RmaIdInputTextbox.addEventListener('keydown', function(event) {
                if (event.keyCode === 13) { // Check if Enter key is pressed
                    const OutboundRmaId = RmaIdInputTextbox.value.trim();
                    console.log('Entered RMA ID:', OutboundRmaId);

                    // Check if the RMA ID input is empty
                    if (OutboundRmaId === '') {
                        // Display error message if RMA ID is empty
                        console.log('⚠️ Please enter RMA ID!');
                        RmaErrorBanner('⚠️ Please enter <strong>RMA ID</strong>!');
                        return;
                    }

                    // Find all checkbox elements by searching for their title attribute
                    const checkboxElements = document.querySelectorAll('input[type="checkbox"][title="Select checkbox to mark this part for Bulk Edit"]');
                    const checkboxesChecked = Array.from(checkboxElements).filter(checkbox => checkbox.checked);
                    const checkboxesToChange = checkboxesChecked.length;

                    if (checkboxesChecked.length === 0) {
                        // Display error message if no row has been selected
                        console.log('⚠️ Please select parts☑️ to change!');
                        RmaErrorBanner('⚠️ Please select parts to change!');
                        return;
                    }

                    // Log the number of selected rows
                    console.log('Number of parts selected to change:', checkboxesToChange);

                    // Select RMA ID checkbox
                    const OutboundRmaIdCheckbox = document.getElementById('edit_outbound_rma_id_checkbox');
                    if (OutboundRmaIdCheckbox) {
                        OutboundRmaIdCheckbox.checked = true;
                        console.log('Outbound RMA ID checkbox checked');

                        // Append text to the RMA ID textbox
                        const OutboundRmaIdTextbox = document.getElementById('outbound_rma_id_textbox');
                        if (OutboundRmaIdTextbox) {
                            // Append the entered RMA ID to the RMA ID textbox
                            OutboundRmaIdTextbox.value = OutboundRmaId;
                            console.log('Outbound RMA ID Textbox value set:', OutboundRmaId);

                            // Set the state dropdown value initially to 36
                            const stateDropdown = document.getElementById('state_state_id');
                            if (stateDropdown) {
                                stateDropdown.value = '36'; // Initially set to 36
                                console.log('State dropdown value set to 36');
                            }

                            // Trigger the "Make Changes" button using XPath
                            const container = document.getElementById('part_search_bulk_edit_popup_dialog_form');
                            if (container) {
                                const makeChangesButtonXPath = `//div[@class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix"]//button[contains(text(), "Make Changes")]`;
                                const makeChangesButton = document.evaluate(makeChangesButtonXPath, container, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

                                if (makeChangesButton) {
                                    makeChangesButton.click();
                                    console.log('Make Changes button clicked');

                                    // Function to periodically check the background color of selected checkbox elements
                                    const checkBackgroundColorRma = setInterval(() => {
                                        let greenCheckboxes = 0;
                                        checkboxElements.forEach((checkboxElement) => {
                                            // Get the parent td element
                                            const tdElement = checkboxElement.closest('td');
                                            // Log the background color of the td element
                                            if (tdElement && window.getComputedStyle(tdElement).getPropertyValue('background-color') === 'rgb(50, 205, 50)') {
                                                greenCheckboxes++;
                                            }
                                        });

                                        if (greenCheckboxes === checkboxesToChange) {
                                            // All checkboxes have turned green, display success message
                                            RmaSuccessBanner(`☑️ x<strong>${checkboxesToChange}</strong> part(S) RMA Processed under #<strong>${OutboundRmaId}</strong>`);
                                            clearInterval(checkBackgroundColorRma);

                                            // Reset cell colors
                                            checkboxElements.forEach((checkboxElement) => {
                                                const tdElement = checkboxElement.closest('td');
                                                if (tdElement) {
                                                    tdElement.style.backgroundColor = 'inherit';
                                                }
                                            });

                                            // Uncheck all selected checkboxes
                                            checkboxesChecked.forEach((checkboxElement) => {
                                                checkboxElement.checked = false;
                                            });

                                            // Uncheck outbound RMA ID checkbox
                                            OutboundRmaIdCheckbox.checked = false;
                                            console.log('Outbound RMA ID checkbox unchecked');

                                            // Reset state dropdown value to 0
                                            stateDropdown.value = '0';
                                            console.log('State dropdown value reset to 0');

                                            // Clear the input textbox
                                            RmaIdInputTextbox.value = '';
                                            // Uncheck the "Bulk Edit Select All" checkbox if it was selected before running the function
                                            const bulkEditSelectAllCheckbox = document.getElementById('part_search_bulk_edit_select_all');
                                            if (bulkEditSelectAllCheckbox && bulkEditSelectAllCheckbox.checked) {
                                                bulkEditSelectAllCheckbox.checked = false;
                                                console.log('Bulk edit select all checkbox unchecked');
                                            }
                                        }
                                    }, 500); // Check every 0.5 seconds
                                } else {
                                    console.log('Make Changes button not found!');
                                }
                            } else {
                                console.log('Container element not found!');
                            }
                        } else {
                            console.log('Outbound RMA ID Textbox not found!');
                        }
                    } else {
                        console.log('Outbound RMA ID checkbox not found!');
                    }
                }
            });
        }



        // Call the function to listen for changes and "Enter" key press
        changeOutboundRmaId();



    }

    // Call the function to create the input textbox and listen for changes and "Enter" key press
    RmaIdInput();

    // Append the input textbox to the document body
    document.body.appendChild(RmaIdInputTextbox);









    /////////////❌ SET GONE  ///////////////////////// *Time/Step reduction: 7 steps to 4 steps.


    // Define the input textbox globally
    const GoneInputTextbox = document.createElement('input');

    function OutboundTrackingIdInput() {
        GoneInputTextbox.type = 'text';
        GoneInputTextbox.placeholder = 'GONE ❌';
        GoneInputTextbox.style.textAlign = 'center'; // Center text




        // Add hover description
        GoneInputTextbox.title = '⚠️ Enter Outbound Tracking ID to set selected ☑️ SNs to GONE';

        // Function to show the yellow banner with moving truck emoji for 3 seconds
        const GoneSuccessBanner = (GoneMessage) => {
            console.log('Show success banner:', GoneMessage);

            // Create the GONE Success banner element
            const GoneBanner = document.createElement('div');
            GoneBanner.innerHTML = GoneMessage; // Use innerHTML instead of textContent
            GoneBanner.style.position = 'fixed';
            GoneBanner.style.top = '0';
            GoneBanner.style.left = '0';
            GoneBanner.style.width = '100%';
            GoneBanner.style.fontSize = '20px';
            GoneBanner.style.backgroundColor = '#e5e7e9';
            GoneBanner.style.color = 'black';
            GoneBanner.style.textAlign = 'center';
            GoneBanner.style.padding = '10px';
            GoneBanner.style.zIndex = '9999';
            GoneBanner.style.border = '1px solid black'; // Add black border

            // Append the banner to the document body
            document.body.appendChild(GoneBanner);

            // After 3 seconds, remove the banner
            setTimeout(() => {
                document.body.removeChild(GoneBanner);
            }, 5000);
        };

        // Function to show the red error banner for 6 seconds
        const GoneErrorBanner = (errorMessage) => {
            console.log('Show error banner:', errorMessage);

            // Create the red banner element
            const GoneErrorBanner = document.createElement('div');
            GoneErrorBanner.innerHTML = errorMessage;
            GoneErrorBanner.style.position = 'fixed';
            GoneErrorBanner.style.top = '0';
            GoneErrorBanner.style.left = '0';
            GoneErrorBanner.style.width = '100%';
            GoneErrorBanner.style.fontSize = '20px';
            GoneErrorBanner.style.backgroundColor = 'red';
            GoneErrorBanner.style.color = 'white';
            GoneErrorBanner.style.textAlign = 'center';
            GoneErrorBanner.style.padding = '10px';
            GoneErrorBanner.style.zIndex = '9999';
            GoneErrorBanner.style.border = '1px solid black'; // Add black border

            // Append the banner to the document body
            document.body.appendChild(GoneErrorBanner);

            // After 6 seconds, remove the banner
            setTimeout(() => {
                document.body.removeChild(GoneErrorBanner);
            }, 3000);
        };

        function changeOutboundTrackingId() {
            GoneInputTextbox.addEventListener('keydown', function(event) {
                if (event.keyCode === 13) { // Check if Enter key is pressed
                    const OutboundTrackingId = GoneInputTextbox.value.trim();
                    console.log('Entered RMA ID:', OutboundTrackingId);

                    // Check if the Outbound Tracking ID input is empty
                    if (OutboundTrackingId === '') {
                        // Display error message if Outbound Tracking ID is empty
                        console.log('⚠️ Please enter Outbound Tracking ID!');
                        GoneErrorBanner('⚠️ Please enter <strong>Outbound Tracking ID</strong>!');
                        return;
                    }

                    console.log('Entered tracking ID:', OutboundTrackingId);

                    // Find all checkbox elements by searching for their title attribute
                    const checkboxElements = document.querySelectorAll('input[type="checkbox"][title="Select checkbox to mark this part for Bulk Edit"]');
                    const checkboxesChecked = Array.from(checkboxElements).filter(checkbox => checkbox.checked);
                    const checkboxesToChange = checkboxesChecked.length;

                    if (checkboxesChecked.length === 0) {
                        // Display error message if no row has been selected
                        console.log('⚠️ Please select parts to change!');
                        GoneErrorBanner('⚠️ Please select parts☑️ to change!');
                        return;
                    }

                    // Log the number of selected rows
                    console.log('Number of rows to change:', checkboxesChecked.length);



                    // Select tracking ID checkbox
                    const OutboundTrackingIdCheckbox = document.getElementById('edit_outbound_tracking_id_checkbox');
                    if (OutboundTrackingIdCheckbox) {
                        OutboundTrackingIdCheckbox.checked = true;
                        console.log('Outbound Tracking ID checkbox checked');

                        // Append text to the tracking ID textbox
                        const OutboundTrackingIdTextbox = document.getElementById('outbound_tracking_id_textbox');
                        if (OutboundTrackingIdTextbox) {
                            // Append the entered tracking ID to the tracking ID textbox
                            OutboundTrackingIdTextbox.value = OutboundTrackingId;
                            console.log('Outbound Tracking ID Textbox value set:', OutboundTrackingId);

                            // Reset the state dropdown value to 16
                            const stateDropdown = document.getElementById('state_state_id');
                            if (stateDropdown) {
                                stateDropdown.value = '16';
                                console.log('State dropdown value set to RMA');
                            }

                            // Trigger the "Make Changes" button using XPath
                            const container = document.getElementById('part_search_bulk_edit_popup_dialog_form');
                            if (container) {
                                const makeChangesButtonXPath = `//div[@class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix"]//button[contains(text(), "Make Changes")]`;
                                const makeChangesButton = document.evaluate(makeChangesButtonXPath, container, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

                                if (makeChangesButton) {
                                    makeChangesButton.click();
                                    console.log('Make Changes button clicked');

                                    // Find all checkbox elements by searching for their title attribute
                                    const checkboxElements = document.querySelectorAll('input[type="checkbox"][title="Select checkbox to mark this part for Bulk Edit"]');
                                    let checkboxesToChange = checkboxElements.length;

                                    // Function to periodically check the background color of selected checkbox elements
                                    const checkBackgroundColorGone = setInterval(() => {
                                        let redCheckboxes = 0;
                                        let greenCheckboxes = 0;
                                        checkboxElements.forEach((checkboxElement) => {
                                            // Get the parent td element
                                            const tdElement = checkboxElement.closest('td');
                                            // Log the background color of the td element
                                            if (tdElement && window.getComputedStyle(tdElement).getPropertyValue('background-color') === 'rgb(50, 205, 50)') {
                                                greenCheckboxes++;
                                            } else if (tdElement && window.getComputedStyle(tdElement).getPropertyValue('background-color') === 'rgb(255, 0, 0)') {
                                                redCheckboxes++;
                                            }
                                        });

                                        if (greenCheckboxes === checkboxElements.length) {
                                            // All checkboxes have turned green, display success message
                                            GoneSuccessBanner(`x<strong>${checkboxesChecked.length}</strong> part(s) set to <strong>GONE ⛔</strong> under <strong>${OutboundTrackingId}</strong>`);
                                            clearInterval(checkBackgroundColorGone);


                                            // Uncheck Outbound Tracking ID checkbox
                                            OutboundTrackingIdCheckbox.checked = false;
                                            console.log('Outbound Tracking ID checkbox unchecked');

                                            // Clear Outbound Tracking ID textbox
                                            OutboundTrackingIdTextbox.value = '';
                                            console.log('Outbound Tracking ID Textbox cleared');

                                            // Reset state dropdown value to 0
                                            stateDropdown.value = '0';
                                            console.log('State dropdown value reset to 0');

                                            // Clear the input textbox
                                            GoneInputTextbox.value = '';
                                            // Uncheck the "Bulk Edit Select All" checkbox if it was selected before running the function
                                            const bulkEditSelectAllCheckbox = document.getElementById('part_search_bulk_edit_select_all');
                                            if (bulkEditSelectAllCheckbox && bulkEditSelectAllCheckbox.checked) {
                                                bulkEditSelectAllCheckbox.checked = false;
                                                console.log('Bulk edit select all checkbox unchecked');
                                            }
                                        } else if (redCheckboxes > 0 && redCheckboxes + greenCheckboxes === checkboxesChecked.length) {
                                            // Display error message if any cell is red and all cells have changed color
                                            GoneErrorBanner(`⚠️ Error: Please ensure part is in <strong>RMA_PROCESSED</strong> state and NOT in user_custody before setting to GONE`);
                                            clearInterval(checkBackgroundColorGone);
                                        }
                                    }, 500); // Check every 0.5 seconds
                                } else {
                                    console.log('Make Changes button not found!');
                                }
                            } else {
                                console.log('Container element not found!');
                            }
                        } else {
                            console.log('Outbound Tracking ID Textbox not found!');
                        }
                    } else {
                        console.log('Outbound Tracking ID checkbox not found!');
                    }
                }
            });
        }

        // Call the function to listen for changes and "Enter" key press
        changeOutboundTrackingId();

    }

    // Call the function to create the input textbox and listen for changes and "Enter" key press
    OutboundTrackingIdInput();

    // Append the input textbox to the document body
    document.body.appendChild(GoneInputTextbox);







    ////🛠️🔄 CHANGE MODEL //////

    // Define modelChangeDropdown globally
    let modelChangeDropdown;

    function showGreenBanner(message) {
        // Check if there's already an existing banner
        const existingBanner = document.getElementById('successBanner');
        if (existingBanner) {
            existingBanner.innerHTML = message; // Update the message using innerHTML
        } else {
            // Create the green banner element
            const modelChangeSuccessBanner = document.createElement('div');
            modelChangeSuccessBanner.id = 'successBanner';
            modelChangeSuccessBanner.innerHTML = message; // Set the message using innerHTML
            modelChangeSuccessBanner.style.position = 'fixed';
            modelChangeSuccessBanner.style.top = '0';
            modelChangeSuccessBanner.style.left = '0';
            modelChangeSuccessBanner.style.width = '100%';
            modelChangeSuccessBanner.style.fontSize = '20px';
            modelChangeSuccessBanner.style.backgroundColor = '#e5e7e9'; // Green background color
            modelChangeSuccessBanner.style.color = 'black';
            modelChangeSuccessBanner.style.textAlign = 'center';
            modelChangeSuccessBanner.style.padding = '10px';
            modelChangeSuccessBanner.style.border = '1px solid black'; // Add black border
            modelChangeSuccessBanner.style.zIndex = '9999';

            // Append the banner to the document body
            document.body.appendChild(modelChangeSuccessBanner);

            // After 4 seconds, remove the banner
            setTimeout(function () {
                document.body.removeChild(modelChangeSuccessBanner);
            }, 4000);
        }
    }

    function showErrorBanner(message) {
        const errorBanner = document.createElement('div');
        errorBanner.id = 'errorBanner';
        errorBanner.innerHTML = message;
        errorBanner.style.position = 'fixed';
        errorBanner.style.top = '0';
        errorBanner.style.left = '0';
        errorBanner.style.width = '100%';
        errorBanner.style.fontSize = '20px';
        errorBanner.style.backgroundColor = '#ffcccc'; // Red background color for error
        errorBanner.style.color = 'black';
        errorBanner.style.textAlign = 'center';
        errorBanner.style.padding = '10px';
        errorBanner.style.border = '1px solid black'; // Add black border
        errorBanner.style.zIndex = '9999';

        document.body.appendChild(errorBanner);

        // After 4 seconds, remove the banner
        setTimeout(function () {
            document.body.removeChild(errorBanner);
        }, 4000);
    }

    function modelChangeMenu() {
        const modelChangeContainer = document.createElement('div');
        modelChangeContainer.style.position = 'fixed';
        modelChangeContainer.style.top = '50%';
        modelChangeContainer.style.left = '50%';
        modelChangeContainer.style.transform = 'translate(-50%, -50%)';
        modelChangeContainer.style.zIndex = '9999';
        modelChangeContainer.style.display = 'flex';
        modelChangeContainer.style.flexDirection = 'column';
        modelChangeContainer.style.alignItems = 'center';

        modelChangeDropdown = document.createElement('select');
        modelChangeDropdown.style.width = '100%';
        modelChangeDropdown.style.boxSizing = 'border-box';
        modelChangeDropdown.style.marginBottom = '10px';

        const defaultOption = document.createElement('option');
        defaultOption.textContent = '🔄Model';
        defaultOption.value = '';
        defaultOption.disabled = true;
        defaultOption.selected = true;
        modelChangeDropdown.appendChild(defaultOption);

        // Fetch available models from the existing dropdown in the popup
        const existingModelDropdown = document.querySelector('select[name="model[model_id]"]');
        if (!existingModelDropdown) {
            console.error('MODEL CHANGE ERROR: Existing model dropdown not found');
            return;
        }

        // Clone the options and append them to the new dropdown
        const options = Array.from(existingModelDropdown.options);
        options.forEach(option => {
            const dropdownOption = document.createElement('option');
            dropdownOption.value = option.value;
            dropdownOption.textContent = option.textContent;
            modelChangeDropdown.appendChild(dropdownOption);
        });

        modelChangeDropdown.addEventListener('change', () => {
            const selectedModelValue = modelChangeDropdown.value;
            const selectedModelText = modelChangeDropdown.options[modelChangeDropdown.selectedIndex].text;
            console.log('Selected model:', selectedModelText);

            if (selectedModelValue) {
                // Get the number of rows currently displayed in the table
                const table = document.querySelector('table.table-bordered.table-striped.table-condensed');
                const numberOfRows = table.querySelectorAll('tbody tr').length; // Count all rows in the tbody of the table

                const confirmMessage = `Are you sure you want to change all ${numberOfRows} displayed serials(s) to '${selectedModelText}'?`;
                if (confirm(confirmMessage)) {
                    executeModelChange(selectedModelValue, selectedModelText); // Pass selected option label
                    // Update the message to include the selected option label
                    const message = `Changing Model to <strong>${selectedModelText}</strong> 🔄`; // Initial Message to display
                    showGreenBanner(message);
                    modelChangeDropdown.value = ''; // Reset the dropdown
                } else {
                    modelChangeDropdown.value = ''; // Reset the dropdown if the user clicks "Cancel"
                }
            }
        });

        modelChangeContainer.appendChild(modelChangeDropdown);
        document.body.appendChild(modelChangeContainer);

        return modelChangeDropdown;
    }

    function executeModelChange(modelValue, selectedLabel) {
        const selectAllCheckbox = document.getElementById('part_search_bulk_edit_select_all');
        if (selectAllCheckbox) {
            selectAllCheckbox.checked = true;
            console.log('Bulk edit checkbox checked');
        } else {
            console.log('Bulk edit checkbox not found!');
        }

        // Select all checkboxes in the table beneath the "Select All" checkbox
        const table = document.querySelector('table.table-bordered.table-striped.table-condensed');
        if (table) {
            const checkboxes = table.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(checkbox => {
                checkbox.checked = true;
            });
            console.log('Number of rows selected:', checkboxes.length);
        } else {
            console.log('Table not found!');
        }

        const bulkEditButton = document.getElementById('part_search_bulk_edit_popup_button');
        if (bulkEditButton) {
            bulkEditButton.click();
            console.log('Bulk edit button clicked');
        } else {
            console.log('Bulk edit button not found!');
        }

        const modelDropdown = document.getElementById('model_model_id');
        if (modelDropdown) {
            modelDropdown.value = modelValue;
            console.log('Model dropdown value set');
        } else {
            console.log('Model dropdown not found!');
        }

        // Trigger the change event on the model dropdown
        const changeEvent = new Event('change', {
            bubbles: true,
            cancelable: true,
        });
        modelDropdown.dispatchEvent(changeEvent);
        console.log('Model dropdown change event triggered');

        // Look for the container that holds both the model dropdown and the "Make Changes" button
        const container = document.getElementById('part_search_bulk_edit_popup_dialog_form');
        if (container) {
            // Find the "Make Changes" button by its text content
            const makeChangesButtonXPath = `//div[@class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix"]//button[contains(text(), "Make Changes")]`;
            const makeChangesButton = document.evaluate(makeChangesButtonXPath, container, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if (makeChangesButton) {
                console.log('Make Changes button found:', makeChangesButton);
                makeChangesButton.click();
                console.log('Make Changes button clicked');
            } else {
                console.log('Make Changes button not found!');
            }
        } else {
            console.log('Container element not found!');
        }

        // Check if all cells have turned green before displaying the success banner
        const checkGreenCellsInterval = setInterval(() => {
            const bulkEditCheckboxCell = document.querySelector('td[title="Changes Successful"]');
            if (!bulkEditCheckboxCell) {
                console.log('Bulk edit checkbox cell not found!');
                clearInterval(checkGreenCellsInterval);
                return;
            }

            const table = bulkEditCheckboxCell.closest('table');
            if (!table) {
                console.log('Table not found!');
                clearInterval(checkGreenCellsInterval);
                return;
            }

            const allCells = [];
            const rows = table.rows;
            const bulkEditCheckboxCellIndex = bulkEditCheckboxCell.cellIndex;
            for (let i = 0; i < rows.length; i++) {
                const cell = rows[i].cells[bulkEditCheckboxCellIndex];
                if (cell) {
                    allCells.push(cell);
                }
            }

            const greenCells = allCells.filter(cell => cell.style.backgroundColor === 'limegreen');
            console.log('Green cells:', greenCells.length);
            console.log('All cells:', allCells.length);

            if (greenCells.length === allCells.length - 1) { // Exclude the bulk edit checkbox cell
                // All cells have turned green, display the success banner
                clearInterval(checkGreenCellsInterval);
                const message = `<strong>${greenCells.length}</strong> part(s) Model Changed to <strong>${selectedLabel}</strong> ✔️`; // Success Message to display
                showGreenBanner(message);

                // Uncheck the "Select All" checkbox
                const selectAllCheckbox = document.getElementById('part_search_bulk_edit_select_all');
                if (selectAllCheckbox) {
                    selectAllCheckbox.checked = false;
                    console.log('"Select All" checkbox unchecked');
                } else {
                    console.log('"Select All" checkbox not found!');
                }

                // Reset the form
                modelDropdown.value = '';

                // Simulate clicking the "Close" button in the dialogue box
                const closeButtonXPath = '/html/body/div[9]/div[3]/div/button'; // Updated XPath expression
                const closeButton = document.evaluate(closeButtonXPath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (closeButton) {
                    closeButton.click();
                    console.log('Close button clicked');
                } else {
                    console.log('Close button not found!');
                }
            }
        }, 1000);
    }

    // Initialize the model change menu
    modelChangeMenu();









    /////////////🔄 STATE CHANGE BUTTON 🔄///////////////////////// *Time/Step reduction: 12s to 3s / 7 steps to 4 steps





    // Modify the stateChangeMenu function to pass the selected option label
    function stateChangeMenu() {
        const stateChangeDropdown = document.createElement('select');
        stateChangeDropdown.setAttribute('title', '⚠️ Changes State of ALL displayed serials'); // Adding tooltip message

        // Define options
        const options = [
            { label: 'STATE 🔄', stateValue: '', categoryValue: '', backgroundColor: '#ffffff', selectable: false },
            { label: 'PENDING_BUILD 🔨', stateValue: '22', categoryValue: '3', backgroundColor: '#f9f9f9', selectable: true },
            { label: 'RESERVED_FOR_BUILD ®️', stateValue: '15', categoryValue: '3', backgroundColor: '#f9f9f9', selectable: true },
            { label: 'DCO_SPARE 🛠️', stateValue: '2', categoryValue: '2', backgroundColor: '#f1f1f1', selectable: true },
            { label: 'PENDING_RMA 🔜', stateValue: '11', categoryValue: '', backgroundColor: '#f1f1f1', selectable: true },
            { label: 'PENDING_SAN 🧹', stateValue: '43', categoryValue: '', backgroundColor: '#f1f1f1', selectable: true }
        ];

        // Create options and append them to the dropdown menu
        options.forEach(option => {
            const dropdownOption = document.createElement('option');
            dropdownOption.value = option.label;
            dropdownOption.textContent = option.label;
            stateChangeDropdown.appendChild(dropdownOption);
        });

        // Attach event listener to the change event of the dropdown menu
        stateChangeDropdown.addEventListener('change', () => {
            const selectedOption = options.find(option => option.label === stateChangeDropdown.value);
            if (selectedOption && selectedOption.selectable) {
                // Get the number of rows currently displayed in the table
                const table = document.querySelector('table.table-bordered.table-striped.table-condensed');
                const numberOfRows = table.querySelectorAll('tbody tr').length; // Count all rows in the tbody of the table

                const confirmMessage = `Are you sure you want to change all ${numberOfRows} displayed serials(s) to '${selectedOption.label}'?`;
                if (confirm(confirmMessage)) {
                    executeStateChange(selectedOption.stateValue, selectedOption.categoryValue, selectedOption.label); // Pass selected option label
                    // Update the message to include the selected option label
                    const message = `Changing State to <strong>${selectedOption.label}</strong> 🔄`; // Initial Message to display
                    showGreenBanner(message);
                }
                stateChangeDropdown.value = 'STATE 🔄'; // Reset back to the default label
            }
        });

        return stateChangeDropdown;
    }



    function executeStateChange(stateValue, categoryValue, selectedLabel) {
        const selectAllCheckbox = document.getElementById('part_search_bulk_edit_select_all');
        if (selectAllCheckbox) {
            selectAllCheckbox.checked = true;
            console.log('Bulk edit checkbox checked');
        } else {
            console.log('Bulk edit checkbox not found!');
        }

        // Select all checkboxes in the table beneath the "Select All" checkbox
        const table = document.querySelector('table.table-bordered.table-striped.table-condensed');
        if (table) {
            const checkboxes = table.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(checkbox => {
                checkbox.checked = true;
            });
            console.log('Number of rows selected:', checkboxes.length);
        } else {
            console.log('Table not found!');
        }

        const bulkEditButton = document.getElementById('part_search_bulk_edit_popup_button');
        if (bulkEditButton) {
            bulkEditButton.click();
            console.log('Bulk edit button clicked');
        } else {
            console.log('Bulk edit button not found!');
        }

        const stateDropdown = document.getElementById('state_state_id');
        if (stateDropdown) {
            stateDropdown.value = stateValue;
            console.log('State dropdown value set');
        } else {
            console.log('State dropdown not found!');
        }

        const categoryDropdown = document.getElementById('category_category_id');
        if (categoryDropdown) {
            categoryDropdown.value = categoryValue;
            console.log('Category dropdown value set');
        } else {
            console.log('Category dropdown not found!');
        }

        categoryDropdown.title = '⚠️ Bulk state change ALL displayed SNs';

        // Trigger the change event on the state dropdown
        const changeEvent = new Event('change', {
            bubbles: true,
            cancelable: true,
        });
        stateDropdown.dispatchEvent(changeEvent);
        console.log('State dropdown change event triggered');

        // Look for the container that holds both the state dropdown and the "Make Changes" button
        const container = document.getElementById('part_search_bulk_edit_popup_dialog_form');
        if (container) {


            // Find the "Make Changes" button by its text content
            const makeChangesButtonXPath = `//div[@class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix"]//button[contains(text(), "Make Changes")]`;
            const makeChangesButton = document.evaluate(makeChangesButtonXPath, container, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if (makeChangesButton) {
                console.log('Make Changes button found:', makeChangesButton);
                makeChangesButton.click();
                console.log('Make Changes button clicked');
            } else {
                console.log('Make Changes button not found!');
            }
        } else {
            console.log('Container element not found!');
        }

        // Check if all cells have turned green before displaying the success banner
        const checkGreenCellsInterval = setInterval(() => {
            const bulkEditCheckboxCell = document.querySelector('td[title="Changes Successful"]');
            if (!bulkEditCheckboxCell) {
                console.log('Bulk edit checkbox cell not found!');
                clearInterval(checkGreenCellsInterval);
                return;
            }

            const table = bulkEditCheckboxCell.closest('table');
            if (!table) {
                console.log('Table not found!');
                clearInterval(checkGreenCellsInterval);
                return;
            }

            const allCells = [];
            const rows = table.rows;
            const bulkEditCheckboxCellIndex = bulkEditCheckboxCell.cellIndex;
            for (let i = 0; i < rows.length; i++) {
                const cell = rows[i].cells[bulkEditCheckboxCellIndex];
                if (cell) {
                    allCells.push(cell);
                }
            }

            const greenCells = allCells.filter(cell => cell.style.backgroundColor === 'limegreen');
            console.log('Green cells:', greenCells.length);
            console.log('All cells:', allCells.length);

            if (greenCells.length === allCells.length - 1) { // Exclude the bulk edit checkbox cell
                // All cells have turned green, display the success banner
                clearInterval(checkGreenCellsInterval);
                const message = `x<strong>${greenCells.length}</strong> part(s) State Changed to <strong>${selectedLabel}</strong> ✔️`; // Success Message to display
                showGreenBanner(message);



                // Uncheck the "Select All" checkbox
                const selectAllCheckbox = document.getElementById('part_search_bulk_edit_select_all');
                if (selectAllCheckbox) {
                    selectAllCheckbox.checked = false;
                    console.log('"Select All" checkbox unchecked');
                } else {
                    console.log('"Select All" checkbox not found!');
                }


                // Reset everything in the form
                stateDropdown.value = '';
                categoryDropdown.value = '';

                // Simulate clicking the "Close" button in the dialogue box
                const closeButtonXPath = '/html/body/div[9]/div[3]/div/button'; // Updated XPath expression
                const closeButton = document.evaluate(closeButtonXPath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (closeButton) {
                    closeButton.click();
                    console.log('Close button clicked');
                } else {
                    console.log('Close button not found!');
                }


            }
        }, 1000);
    }








    ///////////////🌎 TOA TARGET BIN DROPDOWN/////////////////////////////

    // Function to execute actions based on the selected option from the second dropdown menu
    function executeSecondDropdownActions(selectedOption) {
        console.log('Selected Option:', selectedOption);

        // Define options that trigger the same actions
        const triggerOptions = [
            'ARN1',
            'ARN50', 'ARN51', 'ARN52', 'ARN53', 'ARN54', 'ARN56',
            'BAH52', 'BAH53', 'BAH54',
            'CDG53', 'CDG54', 'CDG55', 'CDG70', 'CDG88', 'CDG92',
            'CPT60', 'CPT61', 'CPT62',
            'DXB52', 'DXB53', 'DXB60', 'DXB61', 'DXB62',
            'LHR40', 'LHR53', 'LHR54', 'LHR55', 'LHR56', 'LHR57', 'LHR59', 'LHR63', 'LHR64', 'LHR65', 'LHR79', 'LHR82', 'LHR94', 'LHR112',
            'MAD51', 'MAD55',
            'MXP51', 'MXP53','MXP60', 'MXP61', 'MXP62', 'MXP63', 'MXP64', 'MXP73', 'MXP74', 'MXP80', 'MXP83', 'MXP86',
            'TLV52', 'TLV55', 'TLV60', 'TLV61', 'TLV62',
            'ZAZ60', 'ZAZ61', 'ZAZ62',
            'ZRH60', 'ZRH61', 'ZRH62', 'ZRH65'
        ];




        // Perform actions based on the selected option
        if (triggerOptions.includes(selectedOption)) {
            // Select the "Select All" checkbox
            const selectAllCheckbox = document.getElementById('part_search_bulk_edit_select_all');
            if (selectAllCheckbox) {
                selectAllCheckbox.checked = true;


                // Simulate a click event to trigger any associated actions
                const selectAllChangeEvent = new Event('click', {
                    bubbles: true,
                    cancelable: true,
                });
                selectAllCheckbox.dispatchEvent(selectAllChangeEvent);
                console.log('Select All checkbox clicked');
            }

            // Select bin asset tag checkbox
            const binAssetTagCheckbox = document.getElementById('edit_bin_checkbox');
            if (binAssetTagCheckbox) {
                binAssetTagCheckbox.checked = true;

                // Simulate a click event to trigger any associated actions
                const binAssetTagChangeEvent = new Event('click', {
                    bubbles: true,
                    cancelable: true,
                });
                binAssetTagCheckbox.dispatchEvent(binAssetTagChangeEvent);
                console.log('Bin Asset Tag checkbox clicked');

                // Add conditional text based on selected option to the bin asset tag textbox
                const binAssetTagTextbox = document.getElementById('bin_asset_tag_textbox'); // Adjust the ID if necessary
                if (binAssetTagTextbox) {
                    let additionalText = '';







                    /////////////////🌎 TOA TARGET BIN PER CLUSTER / SITE ////////////////////////// UPDATE BINS HERE!


                    if (selectedOption === 'ARN50' || selectedOption === 'ARN51' || selectedOption === 'ARN52' ) {
                        additionalText += 'PARTS.TOAINBOUND';
                    }

                    else if (selectedOption === 'ARN53' || selectedOption === 'ARN54' || selectedOption === 'ARN56' || selectedOption === 'ARN1' ) {
                        additionalText += 'PARTS.INCOMINGTRANSFER';
                    }


                    // Default text for site codes (BAH)
                    else if (selectedOption.startsWith('BAH')) {
                        // Default text for site codes (BAH)
                        if (selectedOption === 'BAH53') {
                            additionalText += 'PARTS.INBOUND';
                        } else {
                            additionalText += 'DOCK.INBOUND';
                        }
                    }


                    // Gap
                    else if (selectedOption.startsWith('CDG')) {
                        // Default text for site codes (CDG)
                        if (selectedOption === '') {
                            additionalText += '';
                        } else {
                            additionalText += 'PARTS.INCOMINGTRANSFER';
                        }
                    }

                    // Gap
                    else if (selectedOption.startsWith('CPT')) {
                        // Default text for site codes (CPT)
                        additionalText += 'PARTS.TOAINBOUND';
                        if (selectedOption === '') {
                            additionalText += '';
                        }
                    }

                    // Gap
                    else if (selectedOption.startsWith('DUB')) {
                        // Default text for site codes (DUB)
                        if (selectedOption === '') {
                            additionalText += '';
                        } else {
                            additionalText += 'PARTS.INTERNALTRANSFER';
                        }
                    }

                    // Gap
                    else if (selectedOption.startsWith('DXB')) {
                        // Default text for DXB site codes
                        if (selectedOption === 'DXB52') {
                            additionalText += '0-2.TOAINBOUND_INTERNALTRASFER';
                        } else if (selectedOption === 'DXB53') {
                            additionalText += 'TOA_IN_TRANSIT.INBOUND_INTERNALTRANSFER';
                        } else {
                            additionalText += 'TOA_IN_TRANSIT.INCOMINGTRANSFER';
                        }
                    }


                    // Gap
                    else if (selectedOption.startsWith('FRA')) {
                        // Default text for site codes (FRA)
                        additionalText += '.FRA';
                    }

                    // Gap
                    else if (selectedOption.startsWith('LHR')) {
                        // Default text for site codes (LHR)
                        additionalText += 'PARTS.TOAINBOUND';
                    }


                    // Gap
                    else if (selectedOption.startsWith('MXP')) {
                        // Default text for site codes (ZAZ)
                        additionalText += 'PARTS.TOAINBOUND1';
                    }

                    // Gap
                    else if (selectedOption.startsWith('TLV')) {
                        // Default text for TLV site codes
                        if (selectedOption === 'TLV60' || selectedOption === 'TLV61' || selectedOption === 'TLV62' || selectedOption === '' || selectedOption === '') {
                            additionalText += 'PARTS.TOAAPPROVED';
                        } else {
                            additionalText += 'PARTS.INTERNALTRANSFER';
                        }
                    }

                    // Gap
                    else if (selectedOption.startsWith('ZAZ')) {
                        // Default text for site codes (ZAZ)
                        additionalText += 'PARTS.TOAINBOUND';
                    }

                    // Gap
                    else if (selectedOption.startsWith('MAD')) {
                        // Default text for TLV site codes
                        if (selectedOption === 'MAD51') {
                            additionalText += '2-2.TOAINBOUND';
                        } else {
                            additionalText += 'PARTS.TOAINBOUND';
                        }
                    }

                    // Gap
                    else if (selectedOption.startsWith('ZRH')) {
                        // Default text for ZRH site codes
                        if (selectedOption === 'ZRH60' || selectedOption === 'ZRH65') {
                            additionalText += 'PARTS.TOAINBOUND1';
                        } else if (selectedOption === 'ZRH62') {
                            additionalText += 'PARTS.INCOMINGTOA';
                        } else {
                            additionalText += 'PARTS.INBOUND';
                        }
                    } else {
                        additionalText += 'DEFAULT_TEXT_HERE';
                    }

                    // Set the value of the bin asset tag textbox
                    binAssetTagTextbox.value = `${selectedOption}.${additionalText}`;
                    console.log('Bin Asset Tag Textbox value set:', `${selectedOption}.${additionalText}`);




                } else {
                    console.log('Bin Asset Tag Textbox not found!');
                }




                // Trigger the "Make Changes" button
                const container = document.getElementById('part_search_bulk_edit_popup_dialog_form');
                if (container) {
                    // Find the "Make Changes" button by its text content
                    const makeChangesButtonXPath = `//div[@class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix"]//button[contains(text(), "Make Changes")]`;
                    const makeChangesButton = document.evaluate(makeChangesButtonXPath, container, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    if (makeChangesButton) {
                        makeChangesButton.click();
                        console.log('Make Changes button clicked');
                    } else {
                        console.log('Make Changes button not found!');
                    }
                } else {
                    console.log('Container element not found!');
                }
            }
        }
    }





    ///////////////////////🌎 TOA SITE DROPDOWN MENU /////////////////////////////////////////


    function TargetBinTOA() {
        const dropdownMenu2 = document.createElement('select');
        dropdownMenu2.setAttribute('title', '⚠️Bulk tranfser ALL displayed SNs to target Cluster/Site bin'); // Adding tooltip message

        // Define all options
        const options = [
            { label: 'TOA 🌎🗑️', value: '', selectable: false },
            { label: 'ARN50 🇸🇪', value: 'ARN50', selectable: true },
            { label: 'ARN51 🇸🇪', value: 'ARN51', selectable: true },
            { label: 'ARN52 🇸🇪', value: 'ARN52', selectable: true },
            { label: 'ARN53 🇸🇪', value: 'ARN53', selectable: true },
            { label: 'ARN54 🇸🇪', value: 'ARN54', selectable: true },
            { label: 'ARN56 🇸🇪', value: 'ARN56', selectable: true },
            { label: 'BAH52 🇧🇭', value: 'BAH52', selectable: true },
            { label: 'BAH53 🇧🇭', value: 'BAH53', selectable: true },
            { label: 'BAH54 🇧🇭', value: 'BAH54', selectable: true },
            { label: 'CDG53 🇨🇵', value: 'CDG53', selectable: true },
            { label: 'CDG54 🇨🇵', value: 'CDG54', selectable: true },
            { label: 'CDG55 🇨🇵', value: 'CDG55', selectable: true },
            { label: 'CDG70 🇨🇵', value: 'CDG70', selectable: true },
            { label: 'CDG88 🇨🇵', value: 'CDG88', selectable: true },
            { label: 'CDG92 🇨🇵', value: 'CDG92', selectable: true },
            { label: 'CPT60 🇿🇦', value: 'CPT60', selectable: true },
            { label: 'CPT61 🇿🇦', value: 'CPT61', selectable: true },
            { label: 'CPT62 🇿🇦', value: 'CPT62', selectable: true },
            { label: 'DXB52 🇦🇪', value: 'DXB52', selectable: true },
            { label: 'DXB53 🇦🇪', value: 'DXB53', selectable: true },
            { label: 'DXB60 🇦🇪', value: 'DXB60', selectable: true },
            { label: 'DXB61 🇦🇪', value: 'DXB61', selectable: true },
            { label: 'DXB62 🇦🇪', value: 'DXB62', selectable: true },
            { label: 'LHR40 🇬🇧', value: 'LHR40', selectable: true },
            { label: 'LHR53 🇬🇧', value: 'LHR53', selectable: true },
            { label: 'LHR54 🇬🇧', value: 'LHR54', selectable: true },
            { label: 'LHR55 🇬🇧', value: 'LHR55', selectable: true },
            { label: 'LHR56 🇬🇧', value: 'LHR56', selectable: true },
            { label: 'LHR57 🇬🇧', value: 'LHR57', selectable: true },
            { label: 'LHR59 🇬🇧', value: 'LHR59', selectable: true },
            { label: 'LHR63 🇬🇧', value: 'LHR63', selectable: true },
            { label: 'LHR64 🇬🇧', value: 'LHR64', selectable: true },
            { label: 'LHR65 🇬🇧', value: 'LHR65', selectable: true },
            { label: 'LHR79 🇬🇧', value: 'LHR79', selectable: true },
            { label: 'LHR82 🇬🇧', value: 'LHR82', selectable: true },
            { label: 'LHR94 🇬🇧', value: 'LHR94', selectable: true },
            { label: 'MXP51 🇲🇽', value: 'MXP61', selectable: true },
            { label: 'MXP53 🇲🇽', value: 'MXP62', selectable: true },
            { label: 'MXP60 🇲🇽', value: 'MXP60', selectable: true },
            { label: 'MXP61 🇲🇽', value: 'MXP61', selectable: true },
            { label: 'MXP62 🇲🇽', value: 'MXP62', selectable: true },
            { label: 'MXP63 🇲🇽', value: 'MXP63', selectable: true },
            { label: 'MXP64 🇲🇽', value: 'MXP64', selectable: true },
            { label: 'MXP73 🇲🇽', value: 'MXP73', selectable: true },
            { label: 'MXP74 🇲🇽', value: 'MXP74', selectable: true },
            { label: 'MXP80 🇲🇽', value: 'MXP80', selectable: true },
            { label: 'MXP83 🇲🇽', value: 'MXP83', selectable: true },
            { label: 'MXP86 🇲🇽', value: 'MXP86', selectable: true },
            { label: 'TLV60 🇮🇱', value: 'TLV60', selectable: true },
            { label: 'TLV61 🇮🇱', value: 'TLV61', selectable: true },
            { label: 'TLV62 🇮🇱', value: 'TLV62', selectable: true },
            { label: 'TLV52 🇮🇱', value: 'TLV52', selectable: true },
            { label: 'TLV55 🇮🇱', value: 'TLV55', selectable: true },
            { label: 'ZAZ60 🇪🇸', value: 'ZAZ60', selectable: true },
            { label: 'ZAZ61 🇪🇸', value: 'ZAZ61', selectable: true },
            { label: 'ZAZ62 🇪🇸', value: 'ZAZ62', selectable: true },
            { label: 'MAD51 🇪🇸', value: 'MAD51', selectable: true },
            { label: 'MAD55 🇪🇸', value: 'MAD55', selectable: true },
            { label: 'ZRH60 🇨🇭', value: 'ZRH60', selectable: true },
            { label: 'ZRH61 🇨🇭', value: 'ZRH61', selectable: true },
            { label: 'ZRH62 🇨🇭', value: 'ZRH62', selectable: true },
            { label: 'ZRH65 🇨🇭', value: 'ZRH65', selectable: true }

        ];




        // Function to insert separator option
        function insertSeparator() {
            const separatorOption = document.createElement('option');
            separatorOption.setAttribute('disabled', true);
            separatorOption.textContent = '───────────';
            separatorOption.style.height = '1px'; // Set smaller height
            dropdownMenu2.appendChild(separatorOption);
        }


        // Create options and append them to the dropdown menu
        let lastGroup = '';
        options.forEach(option => {
            const currentGroup = option.label.substring(0, 3);
            if (currentGroup !== lastGroup && lastGroup !== '') {
                insertSeparator();
            }
            const dropdownOption = document.createElement('option');
            dropdownOption.value = option.value;
            dropdownOption.textContent = option.label;
            dropdownOption.style.fontSize = '16px';
            dropdownOption.style.fontWeight = 'bold'; // Set font weight to bold
            dropdownOption.style.fontFamily = 'Arial, sans-serif'; // Set font family to Arial, sans-serif


            // Set conditional background color based on value
            if (option.value.includes('ARN')) {
                dropdownOption.style.backgroundColor = '#FFCCCC'; // Light Red for ARN options
            } else if (option.value.includes('BAH')) {
                dropdownOption.style.backgroundColor = '#FFE4B2'; // Light Orange for BAH options
            } else if (option.value.includes('CDG')) {
                dropdownOption.style.backgroundColor = '#B3FFB3'; // Light Green for CDG options
            } else if (option.value.includes('CPT')) {
                dropdownOption.style.backgroundColor = '#B2CCFF'; // Light Blue for CPT options
            } else if (option.value.includes('DUB')) {
                dropdownOption.style.backgroundColor = '#FFFFCC'; // Light Yellow for DUB options
            } else if (option.value.includes('DXB')) {
                dropdownOption.style.backgroundColor = '#E6E6FA'; // Light Purple for DXB options
            } else if (option.value.includes('FRA')) {
                dropdownOption.style.backgroundColor = '#FFB3FF'; // Light Magenta for FRA options
            } else if (option.value.includes('LHR')) {
                dropdownOption.style.backgroundColor = '#FFFF99'; // Light Gold for LHR options
            } else if (option.value.includes('MXP')) {
                dropdownOption.style.backgroundColor = '#F5DEB3'; // Light Brown for MXP options
            } else if (option.value.includes('TLV')) {
                dropdownOption.style.backgroundColor = '#B2DFDB'; // Light Teal for TLV options
            } else if (option.value.includes('ZAZ')) {
                dropdownOption.style.backgroundColor = '#FFC0CB'; // Light DeepPink for ZAZ options
            } else if (option.value.includes('ZRH')) {
                dropdownOption.style.backgroundColor = '#D2B48C'; // Light Olive for ZRH options
            } else {
                dropdownOption.style.backgroundColor = '#D3D3D3'; // Light Grey for other options
            }

            dropdownMenu2.appendChild(dropdownOption);
            lastGroup = currentGroup;
        });





        // Function to show the success banner with moving checkmark emoji for 3 seconds
        function showSuccessBanner(message, quantity) {
            // Check if there's already an existing banner
            const existingBanner = document.getElementById('successBanner');
            if (existingBanner) {
                existingBanner.innerHTML = message; // Update the message
            } else {
                // Create the success banner element
                const ToaSuccessBanner = document.createElement('div');
                ToaSuccessBanner.id = 'successBanner';
                ToaSuccessBanner.innerHTML = message; // Set the message
                ToaSuccessBanner.style.position = 'fixed';
                ToaSuccessBanner.style.top = '0';
                ToaSuccessBanner.style.left = '0';
                ToaSuccessBanner.style.width = '100%';
                ToaSuccessBanner.style.fontSize = '20px';
                ToaSuccessBanner.style.backgroundColor = '#5dade2'; // Green background color
                ToaSuccessBanner.style.color = 'black';
                ToaSuccessBanner.style.textAlign = 'center';
                ToaSuccessBanner.style.padding = '10px';
                ToaSuccessBanner.style.border = '1px solid black'; // Add black border
                ToaSuccessBanner.style.zIndex = '9999';
                ToaSuccessBanner.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Add box shadow for 3D effect

                document.body.appendChild(ToaSuccessBanner);

                // Append the banner to the document body
                document.body.appendChild(ToaSuccessBanner);

                // Create the checkmark emoji element
                const checkmarkEmoji = document.createElement('span');
                checkmarkEmoji.textContent = '📦';
                checkmarkEmoji.style.position = 'absolute';
                checkmarkEmoji.style.left = '-50px'; // Initial position off the right side
                checkmarkEmoji.style.fontSize = '30px'; // Set font size to 14px
                checkmarkEmoji.style.transition = 'left 5s linear'; // Move from right to left in 3 seconds
                ToaSuccessBanner.appendChild(checkmarkEmoji);

                // Trigger reflow to start animation
                void ToaSuccessBanner.offsetWidth;

                // Move the checkmark emoji across the banner
                checkmarkEmoji.style.left = 'calc(100% - 50px)'; // Adjusted to ensure the checkmark fully moves off-screen

                // After 3 seconds, remove the banner
                setTimeout(function() {
                    document.body.removeChild(ToaSuccessBanner);
                    // Reset/clear elements on success
                    resetElements();
                }, 4000);
            }
        }

        // Declare errorBanner outside the function to make it accessible globally
        let errorBanner;

        function showErrorBanner(message) {
            // Check if there's already an existing banner
            const existingBanner = document.getElementById('errorBanner');
            if (existingBanner) {
                existingBanner.innerHTML = message; // Update the message
            } else {
                // Create the red banner element
                const banner = document.createElement('div');
                banner.id = 'errorBanner';
                banner.innerHTML = message; // Set the HTML content
                banner.style.position = 'fixed';
                banner.style.top = '0';
                banner.style.left = '0';
                banner.style.width = '100%';
                banner.style.fontSize = '20px';
                banner.style.backgroundColor = 'red'; // Red background color
                banner.style.color = 'white';
                banner.style.textAlign = 'center';
                banner.style.padding = '10px';
                banner.style.zIndex = '9999';
                banner.style.border = '1px solid black'; // Add black border
                banner.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Add box shadow for 3D effect

                // Append the banner to the document body
                document.body.appendChild(banner);

                // After 6 seconds, remove the banner and uncheck the bulk edit checkbox
                setTimeout(() => {
                    document.body.removeChild(banner);
                    const bulkEditCheckbox = document.querySelector('input[type="checkbox"][title="Select checkbox to mark this part for Bulk Edit"]');
                    if (bulkEditCheckbox) {
                        bulkEditCheckbox.checked = false;
                    }
                }, 6000);
            }
        }

        dropdownMenu2.addEventListener('change', () => {
            const selectedOption = dropdownMenu2.value;
            if (selectedOption !== '') {
                // Find all checkbox elements by searching for their title attribute
                const allCheckboxElements = document.querySelectorAll('input[type="checkbox"][title="Select checkbox to mark this part for Bulk Edit"]');

                // Log the all checkbox elements to the console
                console.log('All Checkbox Elements:', allCheckboxElements);

                // Check if any checkbox element exists
                const quantity = allCheckboxElements.length;

                if (quantity > 0) {
                    // Show confirmation dialog
                    const userConfirmed = confirm(`Are you sure you want to transfer all ${quantity} serial(s) listed to ${selectedOption}?`);

                    if (userConfirmed) {
                        // Call executeSecondDropdownActions function
                        executeSecondDropdownActions(selectedOption);

                        // Function to check background color and display banner
                        checkGreenAndShowBanner(Array.from(allCheckboxElements), selectedOption);

                        // Uncheck all the checkboxes after the action is completed
                        allCheckboxElements.forEach(checkbox => {
                            checkbox.checked = false;
                        });

                        // Reset back to the default label
                        dropdownMenu2.value = '';
                    } else {
                        // Reset back to the default label if user cancels
                        dropdownMenu2.value = '';
                    }
                } else {
                    // No checkboxes found, show error banner
                    showErrorBanner('Reload Page');
                    dropdownMenu2.value = '';
                }
            }
        });

        function checkGreenAndShowBanner(checkboxElements, selectedOption) {
            // Function to periodically check the background color
            const checkBackgroundColor = setInterval(() => {
                let allGreen = true;
                let redDetected = false;

                checkboxElements.forEach(checkboxElement => {
                    // Get the parent td element
                    const tdElement = checkboxElement.closest('td');

                    // Log the background color of the td element
                    if (tdElement) {
                        console.log('Background Color:', window.getComputedStyle(tdElement).getPropertyValue('background-color'));

                        // Check if the background color is not limegreen
                        if (window.getComputedStyle(tdElement).getPropertyValue('background-color') !== 'rgb(50, 205, 50)') {
                            allGreen = false;
                            if (window.getComputedStyle(tdElement).getPropertyValue('background-color') === 'rgb(255, 0, 0)') {
                                redDetected = true;
                            }
                        }
                    }
                });

                // If all cells are green, display the success banner and uncheck bulk edit checkbox
                if (allGreen) {
                    // Display the success banner
                    const message = `x<strong>${checkboxElements.length}</strong> Part(s) moved to <strong>${selectedOption}${getAdditionalText(selectedOption)}</strong> ${getCountryCode(selectedOption)} TOA 🗑️`;
                    showSuccessBanner(message);

                    // Reset elements
                    resetElements(checkboxElements);

                    // Stop checking the background color
                    clearInterval(checkBackgroundColor);
                } else if (redDetected) {
                    // Display the error banner if any cell is red
                    showErrorBanner(`⚠️ ERROR with ${selectedOption} TOA Target Bin - Please contact <a href="https://phonetool.AMAZON.com/users/pjbyrne" target="_blank">pjbyrne</a>`);
                    // Stop checking the background color
                    clearInterval(checkBackgroundColor);
                }
            }, 500); // Check every 0.5 seconds
        }

        function resetElements(checkboxElements) {
            // Uncheck the "Select All" checkbox
            const selectAllCheckbox = document.getElementById('part_search_bulk_edit_select_all');
            if (selectAllCheckbox) {
                selectAllCheckbox.checked = false;
                console.log('"Select All" checkbox unchecked');
            } else {
                console.log('"Select All" checkbox not found!');
            }

            // Uncheck bin asset tag checkbox
            const binAssetTagCheckbox = document.getElementById('edit_bin_checkbox');
            if (binAssetTagCheckbox) {
                binAssetTagCheckbox.checked = false;
                console.log('Bin Asset Tag checkbox unchecked');
            } else {
                console.log('Bin Asset Tag checkbox not found!');
            }

            // Clear bin asset tag textbox
            const binAssetTagTextbox = document.getElementById('bin_asset_tag_textbox');
            if (binAssetTagTextbox) {
                binAssetTagTextbox.value = '';
                console.log('Bin Asset Tag textbox cleared');
            } else {
                console.log('Bin Asset Tag textbox not found!');
            }

            // Reset the background color of the table cells
            checkboxElements.forEach(checkboxElement => {
                const tdElement = checkboxElement.closest('td');
                if (tdElement) {
                    tdElement.style.backgroundColor = ''; // Reset to original background color
                }
            });
        }

        return dropdownMenu2;











        // Function to get the country code based on the selected option
        function getCountryCode(selectedOption) {
            let countryCode = '';
            if (selectedOption.includes('ARN')) {
                countryCode = '🇸🇪'; // Sweden
            } else if (selectedOption.includes('BAH')) {
                countryCode = '🇧🇭'; // Bahrain
            } else if (selectedOption.includes('CDG')) {
                countryCode = '🇫🇷'; // France
            } else if (selectedOption.includes('CPT')) {
                countryCode = '🇿🇦'; // South Africa
            } else if (selectedOption.includes('DXB')) {
                countryCode = '🇦🇪'; // Dubai
            } else if (selectedOption.includes('LHR')) {
                countryCode = '🇬🇧'; // United Kingdom
            } else if (selectedOption.includes('MAD')) {
                countryCode = '🇪🇸'; // Spain
            } else if (selectedOption.includes('MXP')) {
                countryCode = '🇮🇹'; // Italy
            } else if (selectedOption.includes('TLV')) {
                countryCode = '🇮🇱'; // Israel
            } else if (selectedOption.includes('ZAZ')) {
                countryCode = '🇪🇸'; // Spain
            } else if (selectedOption.includes('ZRH')) {
                countryCode = '🇨🇭'; // Switzerland
            }
            return countryCode;
        }


    }








    //////////////// 🗑️LOCAL SITE TRANSFER TARGET BIN🗑️////////////////// *NEEDS TO BE FIXED/IMPROVED!


    // Function to execute actions based on the selected option from the second dropdown menu
    function executeThirdDropdownActions(selectedOption) {
        console.log('Selected Option:', selectedOption);

        // Define options that trigger the same actions
        const triggerOptions = [
            'ARN1',
            'ARN50', 'ARN51', 'ARN52', 'ARN53', 'ARN54', 'ARN56',
            'BAH52', 'BAH53', 'BAH54',
            'CDG53', 'CDG54', 'CDG55', 'CDG70', 'CDG88', 'CDG92',
            'CPT60', 'CPT61', 'CPT62',
            'DXB52', 'DXB53', 'DXB60', 'DXB61', 'DXB62',
            'LHR40', 'LHR5', 'LHR53', 'LHR54', 'LHR55', 'LHR56', 'LHR57', 'LHR59', 'LHR61', 'LHR63', 'LHR64', 'LHR65', 'LHR79', 'LHR82', 'LHR94', 'LHR112',
            'MAD51', 'MAD55',
            'MXP51', 'MXP53','MXP60', 'MXP61', 'MXP62', 'MXP63', 'MXP64', 'MXP73', 'MXP74', 'MXP80', 'MXP83', 'MXP86',
            'TLV52', 'TLV55', 'TLV60', 'TLV61', 'TLV62',
            'ZAZ60', 'ZAZ61', 'ZAZ62',
            'ZRH60', 'ZRH61', 'ZRH62', 'ZRH65'
        ];






        // Perform actions based on the selected option
        if (triggerOptions.includes(selectedOption)) {
            // Select the "Select All" checkbox
            const selectAllCheckbox = document.getElementById('part_search_bulk_edit_select_all');
            if (selectAllCheckbox) {
                selectAllCheckbox.checked = true;



                // Simulate a click event to trigger any associated actions
                const selectAllChangeEvent = new Event('click', {
                    bubbles: true,
                    cancelable: true,
                });
                selectAllCheckbox.dispatchEvent(selectAllChangeEvent);
                console.log('Select All checkbox clicked');
            }

            // Select bin asset tag checkbox
            const binAssetTagCheckbox = document.getElementById('edit_bin_checkbox');
            if (binAssetTagCheckbox) {
                binAssetTagCheckbox.checked = true;

                // Simulate a click event to trigger any associated actions
                const binAssetTagChangeEvent = new Event('click', {
                    bubbles: true,
                    cancelable: true,
                });
                binAssetTagCheckbox.dispatchEvent(binAssetTagChangeEvent);
                console.log('Bin Asset Tag checkbox clicked');

                // Add conditional text based on selected option to the bin asset tag textbox
                const binAssetTagTextbox = document.getElementById('bin_asset_tag_textbox'); // Adjust the ID if necessary
                if (binAssetTagTextbox) {
                    let additionalText = '';






                    ///////////🗑️LOCAL SITE TRANSFER TARGET BIN PER SITE ////////////////////////// UPDATE BINS HERE!


                    if (selectedOption.includes('ARN')) {
                        // Default text for ARN site codes
                        additionalText += '.PARTS.INCOMINGTRANSFER';
                    }




                    else if (selectedOption.startsWith('BAH')) {


                        // Default text for site codes (BAH) - use this type of code snippet if there are varying bin names
                        if (selectedOption === 'BAH53') {
                            additionalText += '.PARTS.INBOUND';
                        } else {
                            additionalText += '.DOCK.INBOUND';
                        }
                    }


                    // Gap
                    else if (selectedOption.startsWith('CDG')) {
                        // Default text for CDG site codes
                        additionalText += '.PARTS.INBOUND';

                        // Specific text for CDG site codes
                        if (selectedOption === '') {
                            additionalText += '';
                        }
                    }
                    // Gap
                    else if (selectedOption.startsWith('CPT')) {
                        // Default text for CPT site codes
                        additionalText += '.PARTS.INBOUND';

                        // Specific text for CPT site codes
                        if (selectedOption === '') {
                            additionalText += '.SOME_OTHER_TEXT_FOR_CPT61';
                        }
                    }
                    // Gap
                    else if (selectedOption.startsWith('DUB')) {
                        // Default text for DUB site codes
                        additionalText += '.DUB';
                    }
                    // Gap
                    else if (selectedOption.startsWith('DXB')) {
                        // Default text for DXB site codes
                        if (selectedOption === 'DXB52') {
                            additionalText += '.0-2.TOAINBOUND_INTERNALTRASFER';
                        } else if (selectedOption === 'DXB53') {
                            additionalText += '.TOA_IN_TRANSIT.INBOUND_INTERNALTRANSFER';
                        } else {
                            additionalText += '.TOA_IN_TRANSIT.INCOMINGTRANSFER';
                        }
                    }

                    // Gap
                    else if (selectedOption.startsWith('FRA')) {
                        // Default text for FRA site codes
                        additionalText += '.FRA';
                    }
                    // Gap
                    else if (selectedOption.startsWith('LHR')) {
                        // Default text for LHR site codes
                        additionalText += '.PARTS.INBOUND';
                    }
                    // Gap
                    else if (selectedOption.startsWith('MXP')) {
                        // Default text for MXP site codes
                        additionalText += '.PARTS.INTRANSIT';

                        // Specific text for MXP site codes
                        if (selectedOption === '') {
                            additionalText += '.SOME_OTHER_TEXT_FOR_MXP60';
                        }
                    }
                    // Gap
                    else if (selectedOption.startsWith('TLV')) {
                        // Default text for TLV site codes
                        if (selectedOption === 'TLV60' || selectedOption === 'TLV61' || selectedOption === 'TLV62' || selectedOption === '' || selectedOption === '') {
                            additionalText += '.PARTS.TOAAPPROVED';
                        } else {
                            additionalText += '.PARTS.INTERNALTRANSFER';
                        }
                    }
                    // Gap
                    else if (selectedOption.startsWith('ZAZ')) {
                        // Default text for ZAZ site codes
                        additionalText += '.PARTS.INTERNALTRANSFER';
                    }


                    // Gap
                    else if (selectedOption === 'MAD55') {
                        // Default text for MXP site codes
                        additionalText += '.PARTS.INTERNALTRANSFER';


                        if (selectedOption === 'MAD51') {
                            additionalText += '.2-2.INTERNALTRANSFER';
                        }
                    }

                    // Gap
                    else if (selectedOption.startsWith('ZRH')) {
                        // Default text for ZRH site codes
                        additionalText += '.ZRH';
                    }
                    else {
                        // Default text for other site codes
                        additionalText += 'DEFAULT_TEXT_HERE';
                    }

                    // Add more conditions for additional labels as needed

                    // Set the value of the bin asset tag textbox
                    binAssetTagTextbox.value = selectedOption + additionalText;
                    console.log('Bin Asset Tag Textbox value set:', selectedOption + additionalText);



                } else {
                    console.log('Bin Asset Tag Textbox not found!');
                }




                // Trigger the "Make Changes" button
                const container = document.getElementById('part_search_bulk_edit_popup_dialog_form');
                if (container) {
                    // Find the "Make Changes" button by its text content
                    const makeChangesButtonXPath = `//div[@class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix"]//button[contains(text(), "Make Changes")]`;
                    const makeChangesButton = document.evaluate(makeChangesButtonXPath, container, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    if (makeChangesButton) {
                        makeChangesButton.click();
                        console.log('Make Changes button clicked');
                    } else {
                        console.log('Make Changes button not found!');
                    }
                } else {
                    console.log('Container element not found!');
                }
            }
        }
    }




    ///////////////////////🗑️LOCAL SITE TRANSFER DROPDOWN MENU /////////////////////////////////////////


    function SiteTransferBin() {
        const dropdownMenu3 = document.createElement('select');
        dropdownMenu3.setAttribute('title', '⚠️ Bulk transfer ALL displayed SNs to target Site bin'); // Adding tooltip message

        // Retrieve cluster code from the page
        const clusterCodeElement = document.querySelector('.search_refinements .linklist li');
        const clusterCode = clusterCodeElement.textContent.trim(); // Extracting text content and removing leading/trailing whitespace

        // Define all options
        const allOptions = [
            { label: 'LOCAL 📍🗑️', value: '', selectable: false },
            { label: 'ARN50 🇸🇪', value: 'ARN50', selectable: true },
            { label: 'ARN51 🇸🇪', value: 'ARN51', selectable: true },
            { label: 'ARN52 🇸🇪', value: 'ARN52', selectable: true },
            { label: 'ARN53 🇸🇪', value: 'ARN53', selectable: true },
            { label: 'ARN54 🇸🇪', value: 'ARN54', selectable: true },
            { label: 'ARN56 🇸🇪', value: 'ARN56', selectable: true },
            { label: 'BAH52 🇧🇭', value: 'BAH52', selectable: true },
            { label: 'BAH53 🇧🇭', value: 'BAH53', selectable: true },
            { label: 'BAH54 🇧🇭', value: 'BAH54', selectable: true },
            { label: 'CDG53 🇨🇵', value: 'CDG53', selectable: true },
            { label: 'CDG54 🇨🇵', value: 'CDG54', selectable: true },
            { label: 'CDG55 🇨🇵', value: 'CDG55', selectable: true },
            { label: 'CDG70 🇨🇵', value: 'CDG70', selectable: true },
            { label: 'CDG88 🇨🇵', value: 'CDG88', selectable: true },
            { label: 'CDG92 🇨🇵', value: 'CDG92', selectable: true },
            { label: 'CPT60 🇿🇦', value: 'CPT60', selectable: true },
            { label: 'CPT61 🇿🇦', value: 'CPT61', selectable: true },
            { label: 'CPT62 🇿🇦', value: 'CPT62', selectable: true },
            { label: 'DXB52 🇦🇪', value: 'DXB52', selectable: true },
            { label: 'DXB53 🇦🇪', value: 'DXB53', selectable: true },
            { label: 'DXB60 🇦🇪', value: 'DXB60', selectable: true },
            { label: 'DXB61 🇦🇪', value: 'DXB61', selectable: true },
            { label: 'DXB62 🇦🇪', value: 'DXB62', selectable: true },
            { label: 'LHR40 🇬🇧', value: 'LHR40', selectable: true },
            { label: 'LHR5 🇬🇧', value: 'LHR5', selectable: true },
            { label: 'LHR53 🇬🇧', value: 'LHR53', selectable: true },
            { label: 'LHR54 🇬🇧', value: 'LHR54', selectable: true },
            { label: 'LHR55 🇬🇧', value: 'LHR55', selectable: true },
            { label: 'LHR56 🇬🇧', value: 'LHR56', selectable: true },
            { label: 'LHR57 🇬🇧', value: 'LHR57', selectable: true },
            { label: 'LHR59 🇬🇧', value: 'LHR59', selectable: true },
            { label: 'LHR61 🇬🇧', value: 'LHR61', selectable: true },
            { label: 'LHR63 🇬🇧', value: 'LHR63', selectable: true },
            { label: 'LHR64 🇬🇧', value: 'LHR64', selectable: true },
            { label: 'LHR65 🇬🇧', value: 'LHR65', selectable: true },
            { label: 'LHR79 🇬🇧', value: 'LHR79', selectable: true },
            { label: 'LHR82 🇬🇧', value: 'LHR82', selectable: true },
            { label: 'LHR94 🇬🇧', value: 'LHR94', selectable: true },
            { label: 'MXP51 🇲🇽', value: 'MXP61', selectable: true },
            { label: 'MXP53 🇲🇽', value: 'MXP62', selectable: true },
            { label: 'MXP60 🇲🇽', value: 'MXP60', selectable: true },
            { label: 'MXP61 🇲🇽', value: 'MXP61', selectable: true },
            { label: 'MXP62 🇲🇽', value: 'MXP62', selectable: true },
            { label: 'MXP63 🇲🇽', value: 'MXP63', selectable: true },
            { label: 'MXP64 🇲🇽', value: 'MXP64', selectable: true },
            { label: 'MXP73 🇲🇽', value: 'MXP73', selectable: true },
            { label: 'MXP74 🇲🇽', value: 'MXP74', selectable: true },
            { label: 'MXP80 🇲🇽', value: 'MXP80', selectable: true },
            { label: 'MXP83 🇲🇽', value: 'MXP83', selectable: true },
            { label: 'MXP86 🇲🇽', value: 'MXP86', selectable: true },
            { label: 'TLV60 🇮🇱', value: 'TLV60', selectable: true },
            { label: 'TLV61 🇮🇱', value: 'TLV61', selectable: true },
            { label: 'TLV62 🇮🇱', value: 'TLV62', selectable: true },
            { label: 'TLV52 🇮🇱', value: 'TLV52', selectable: true },
            { label: 'TLV55 🇮🇱', value: 'TLV55', selectable: true },
            { label: 'ZAZ60 🇪🇸', value: 'ZAZ60', selectable: true },
            { label: 'ZAZ61 🇪🇸', value: 'ZAZ61', selectable: true },
            { label: 'ZAZ62 🇪🇸', value: 'ZAZ62', selectable: true },
            { label: 'MAD51 🇪🇸', value: 'MAD51', selectable: true },
            { label: 'MAD55 🇪🇸', value: 'MAD55', selectable: true },
            { label: 'ZRH60 🇨🇭', value: 'ZRH60', selectable: true },
            { label: 'ZRH61 🇨🇭', value: 'ZRH61', selectable: true },
            { label: 'ZRH62 🇨🇭', value: 'ZRH62', selectable: true },
            { label: 'ZRH65 🇨🇭', value: 'ZRH65', selectable: true }

        ];




        // Filter options based on the presence of site codes in the cluster code
        const filteredOptions = allOptions.filter(option => clusterCode.includes(option.value.substring(0, 3)));

        // Create options and append them to the dropdown menu
        filteredOptions.forEach(option => {
            const dropdownOption = document.createElement('option');
            dropdownOption.value = option.value;
            dropdownOption.textContent = option.label;
            dropdownOption.style.fontSize = '16px';
            dropdownOption.style.fontWeight = 'bold'; // Set font weight to bold
            dropdownOption.style.fontFamily = 'Arial, sans-serif'; // Set font family to Arial, sans-serif

            // Set conditional background color based on value
            if (option.value.includes('ARN')) {
                dropdownOption.style.backgroundColor = '#FFCCCC'; // Light Red for ARN options
            } else if (option.value.includes('BAH')) {
                dropdownOption.style.backgroundColor = '#FFE4B2'; // Light Orange for BAH options
            } else if (option.value.includes('CDG')) {
                dropdownOption.style.backgroundColor = '#B3FFB3'; // Light Green for CDG options
            } else if (option.value.includes('CPT')) {
                dropdownOption.style.backgroundColor = '#B2CCFF'; // Light Blue for CPT options
            } else if (option.value.includes('DUB')) {
                dropdownOption.style.backgroundColor = '#FFFFCC'; // Light Yellow for DUB options
            } else if (option.value.includes('DXB')) {
                dropdownOption.style.backgroundColor = '#E6E6FA'; // Light Purple for DXB options
            } else if (option.value.includes('FRA')) {
                dropdownOption.style.backgroundColor = '#FFB3FF'; // Light Magenta for FRA options
            } else if (option.value.includes('LHR')) {
                dropdownOption.style.backgroundColor = '#FFFF99'; // Light Gold for LHR options
            } else if (option.value.includes('MXP')) {
                dropdownOption.style.backgroundColor = '#F5DEB3'; // Light Brown for MXP options
            } else if (option.value.includes('TLV')) {
                dropdownOption.style.backgroundColor = '#B2DFDB'; // Light Teal for TLV options
            } else if (option.value.includes('ZAZ')) {
                dropdownOption.style.backgroundColor = '#FFC0CB'; // Light DeepPink for ZAZ options
            } else if (option.value.includes('ZRH')) {
                dropdownOption.style.backgroundColor = '#D2B48C'; // Light Olive for ZRH options
            } else {
                dropdownOption.style.backgroundColor = '#D3D3D3'; // Light Grey for other options
            }

            dropdownMenu3.appendChild(dropdownOption);
        });









        ////// 🗑️LOCAL SITE TRANSFER SUCCESS BANNER ✅ ////////////////////////////

        // Function to show the yellow banner with moving trash can emoji for 3 seconds
        function showYellowBanner(message, quantity) {
            // Check if there's already an existing banner
            const existingBanner = document.getElementById('warningBanner');
            if (existingBanner) {
                existingBanner.innerHTML = message; // Update the message
            } else {
                // Create the yellow banner element
                const banner = document.createElement('div');
                banner.id = 'warningBanner';
                banner.innerHTML = message;
                banner.style.position = 'fixed';
                banner.style.top = '0';
                banner.style.left = '0';
                banner.style.width = '100%';
                banner.style.fontSize = '20px';
                banner.style.backgroundColor = 'yellow'; // Yellow background color
                banner.style.color = 'black';
                banner.style.textAlign = 'center';
                banner.style.fontWeight = 'bold';
                banner.style.padding = '10px';
                banner.style.zIndex = '9999';
                banner.style.border = '1px solid black'; // Add black border
                banner.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Add box shadow for 3D effect

                // Append the banner to the document body
                document.body.appendChild(banner);

                // Create the box emoji element
                const boxEmoji = document.createElement('span');
                boxEmoji.textContent = '📦';
                boxEmoji.style.position = 'absolute';
                boxEmoji.style.left = '-50px'; // Initial position off the left side
                boxEmoji.style.transition = 'left 5s linear'; // Move from left to right in 3 seconds
                boxEmoji.style.fontSize = '30px'; // Set the font size to make the emoji larger
                banner.appendChild(boxEmoji);

                // Trigger reflow to start animation
                void banner.offsetWidth;

                // Move the trash can emoji across the banner
                boxEmoji.style.left = 'calc(100% - 50px)'; // Adjusted to ensure the trash can fully moves off-screen

                // After 3 seconds, remove the banner and uncheck the bulk edit checkbox
                setTimeout(() => {
                    document.body.removeChild(banner);
                    resetElements();
                }, 2500);
            }
        }







        ///////🗑️LOCAL SITE TRANSFER   ERROR BANNER! ////////////////////////////////

        function showErrorBanner(message) {
            // Check if there's already an existing banner
            const existingBanner = document.getElementById('errorBanner');
            if (existingBanner) {
                existingBanner.innerHTML = message; // Update the message
            } else {
                // Create the red banner element
                const banner = document.createElement('div');
                banner.id = 'errorBanner';
                banner.innerHTML = message; // Set the HTML content
                banner.style.position = 'fixed';
                banner.style.top = '0';
                banner.style.left = '0';
                banner.style.width = '100%';
                banner.style.fontSize = '20px';
                banner.style.backgroundColor = 'red'; // Red background color
                banner.style.color = 'white';
                banner.style.textAlign = 'center';
                banner.style.padding = '10px';
                banner.style.zIndex = '9999';
                banner.style.border = '1px solid black'; // Add black border
                banner.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Add box shadow for 3D effect

                // Append the banner to the document body
                document.body.appendChild(banner);

                // After 3 seconds, remove the banner and uncheck the bulk edit checkbox
                setTimeout(() => {
                    document.body.removeChild(banner);
                    const bulkEditCheckbox = document.querySelector('input[type="checkbox"][title="Select checkbox to mark this part for Bulk Edit"]');
                    if (bulkEditCheckbox) {
                        bulkEditCheckbox.checked = false;
                    }
                }, 6000);
            }
        }

        dropdownMenu3.addEventListener('change', () => {
            const selectedOption = dropdownMenu3.value;
            if (selectedOption !== '') {
                // Find all checkbox elements by searching for their title attribute
                const allCheckboxElements = document.querySelectorAll('input[type="checkbox"][title="Select checkbox to mark this part for Bulk Edit"]');

                // Log the all checkbox elements to the console
                console.log('All Checkbox Elements:', allCheckboxElements);

                // Check if any checkbox element exists
                const quantity = allCheckboxElements.length;

                if (quantity > 0) {
                    // Show confirmation dialog
                    const userConfirmed = confirm(`Are you sure you want to transfer all ${quantity} serial(s) listed to ${selectedOption}?`);

                    if (userConfirmed) {
                        // Call executeThirdDropdownActions function
                        executeThirdDropdownActions(selectedOption);

                        // Function to check background color and display banner
                        checkGreenAndShowBanner(Array.from(allCheckboxElements), selectedOption);

                        // Uncheck all the checkboxes after the action is completed
                        allCheckboxElements.forEach(checkbox => {
                            checkbox.checked = false;
                        });

                        // Reset back to the default label
                        dropdownMenu3.value = '';
                    } else {
                        // Reset back to the default label if user cancels
                        dropdownMenu3.value = '';
                    }
                } else {
                    // No checkboxes found, show error banner
                    showErrorBanner('Please reload page.');
                    dropdownMenu3.value = '';
                }
            }
        });

        // Function to check if all checkboxes have turned green and then display the success banner
        function checkGreenAndShowBanner(checkboxElements, selectedOption) {
            // Function to periodically check the background color
            const checkBackgroundColor = setInterval(() => {
                let allGreen = true;
                let redDetected = false;

                checkboxElements.forEach(checkboxElement => {
                    // Get the parent td element
                    const tdElement = checkboxElement.closest('td');

                    // Log the background color of the td element
                    if (tdElement) {
                        console.log('Background Color:', window.getComputedStyle(tdElement).getPropertyValue('background-color'));

                        // Check if the background color is not limegreen
                        if (window.getComputedStyle(tdElement).getPropertyValue('background-color') !== 'rgb(50, 205, 50)') {
                            allGreen = false;
                            if (window.getComputedStyle(tdElement).getPropertyValue('background-color') === 'rgb(255, 0, 0)') {
                                redDetected = true;
                            }
                        }
                    }
                });

                // If all cells are green, display the success banner and uncheck bulk edit checkbox
                if (allGreen) {
                    // Display the success banner
                    const message = `x<strong>${checkboxElements.length}</strong> Part(s) transferred to <strong>${selectedOption}</strong> 🗑️`;
                    showYellowBanner(message, checkboxElements.length);

                    // Reset elements
                    resetElements(checkboxElements);

                    // CLICK CLOSE BUTTON
                    const closeButtonXPaths = [
                        '/html/body/div[7]/div[3]/div/button',
                        '/html/body/div[8]/div[3]/div/button',
                        '/html/body/div[9]/div[3]/div/button'
                    ];

                    // Initialize a variable to track if the Close button is found
                    let closeButtonFound = false;

                    // Iterate through the XPath options to find and click the Close button
                    for (const xpath of closeButtonXPaths) {
                        const closeButton = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                        if (closeButton) {
                            closeButton.click();
                            console.log('Close button clicked');
                            closeButtonFound = true;
                            break; // Exit the loop if Close button is found and clicked
                        }
                    }

                    // Log a message if the Close button is not found
                    if (!closeButtonFound) {
                        console.log('Close button not found!');
                    }

                    // Stop checking the background color
                    clearInterval(checkBackgroundColor);
                } else if (redDetected) {
                    // Display the error banner if any cell is red
                    const errorMessage = `⚠️ ERROR with <strong>${selectedOption}</strong> Local Target Bin - Please contact <a href="https://phonetool.AMAZON.com/users/pjbyrne" target="_blank"><strong>pjbyrne</strong>`;
                    showErrorBanner(errorMessage);

                    // Stop checking the background color
                    clearInterval(checkBackgroundColor);
                }
            }, 500); // Check every 0.5 seconds
        }

        // Function to reset/clear elements
        function resetElements(checkboxElements) {
            // Uncheck the "Select All" checkbox
            const selectAllCheckbox = document.getElementById('part_search_bulk_edit_select_all');
            if (selectAllCheckbox) {
                selectAllCheckbox.checked = false;
                console.log('"Select All" checkbox unchecked');
            } else {
                console.log('"Select All" checkbox not found!');
            }

            // Uncheck bin asset tag checkbox
            const binAssetTagCheckbox = document.getElementById('edit_bin_checkbox');
            if (binAssetTagCheckbox) {
                binAssetTagCheckbox.checked = false;
                console.log('Bin Asset Tag checkbox unchecked');
            } else {
                console.log('Bin Asset Tag checkbox not found!');
            }

            // Clear bin asset tag textbox
            const binAssetTagTextbox = document.getElementById('bin_asset_tag_textbox');
            if (binAssetTagTextbox) {
                binAssetTagTextbox.value = '';
                console.log('Bin Asset Tag textbox cleared');
            } else {
                console.log('Bin Asset Tag textbox not found!');
            }

            // Reset the background color of the table cells
            checkboxElements.forEach(checkboxElement => {
                const tdElement = checkboxElement.closest('td');
                if (tdElement) {
                    tdElement.style.backgroundColor = ''; // Reset to original background color
                }
            });
        }

        return dropdownMenu3;







    }




    ///// GET CLUSTER NAME FOR LOCAL TRANSFER OPTIONS ////////
    function getClusterName(clusterCode) {
        const siteCode = clusterCode.substring(0, 3); // Extract site code from cluster code
        // Logic to determine the cluster name based on the site code
        // Example: if siteCode is BAH, return 'BAH'
        // You can extend this logic for other site codes as needed
        return siteCode;
    }

    ////////// GET BIN DATA FOR SUCCESS BANNER /////////
    function getAdditionalText(selectedOption) {
        let additionalText = '';


        // Add more conditions for additional bins as needed
        return additionalText;
    }









    ///// FETCH SAN REQUIREMENT /////

    // Function to fetch sanitization requirement data step by step
    const fetchSanRequirementStepByStep = async () => {
        try {
            // Select the table rows
            const rows = document.querySelectorAll('table.table-bordered.table-striped.table-condensed tbody tr');
            if (!rows.length) {
                console.error('Error: No table rows found');
                return;
            }

            // Initialize counters and serials for each sanitization requirement category
            let greenCount = 0;
            let yellowCount = 0;
            let redCount = 0;
            let greenSerials = [];
            let yellowSerials = [];
            let redSerials = [];

            // Loop through each row and fetch the sanitization requirement data
            rows.forEach(row => {
                // Get the partId from the second column of each row and trim it
                const partId = row.querySelectorAll('td')[1].textContent.trim().replace(/\s*📦$/, ''); // Trim and remove emoji
                const serialId = row.querySelectorAll('td')[2].textContent.trim().replace(/\s*\(\s*RMA\s*\)$/i, ''); // Get the serial ID from the serial cell and clean it up
                const serialCell = row.querySelectorAll('td')[2].style;

                // Construct the URL with the partId
                const apiUrl = 'https://mobility.AMAZON.com/part/part/';
                const corsUrl = apiUrl + encodeURIComponent(partId);

                console.log('SanReq Fetching URL:', corsUrl);

                // Make a GET request using GM_xmlhttpRequest
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: corsUrl,
                    onload: function (response) {
                        if (response.status === 200) {
                            // Create a temporary div to parse the HTML response
                            const tempDiv = document.createElement('div');
                            tempDiv.innerHTML = response.responseText;

                            // Find the sanitization requirement element and extract its text
                            const sanitizationRequirementElement = tempDiv.querySelector('.sanitization_requirement_div');
                            if (sanitizationRequirementElement) {
                                const sanitizationRequirement = sanitizationRequirementElement.textContent.trim();
                                console.log('Sanitization Requirement:', sanitizationRequirement);

                                // Count the rows based on sanitization requirement category
                                if (sanitizationRequirement === 'Not Required') {
                                    greenCount++;
                                    greenSerials.push(serialId);
                                    serialCell.borderRight = '3px solid #2ecc71'; // Add right border to serial cell
                                    serialCell.borderBottom = '1px solid black'; // Add bottom border
                                    serialCell.borderTop = '1px solid black'; // Add top border
                                    serialCell.borderLeft = '3px solid #2ecc71'; // Add GREEN left border
                                } else if (sanitizationRequirement === 'Required') {
                                    yellowCount++;
                                    yellowSerials.push(serialId);
                                    serialCell.borderRight = '3px solid #f1c40f'; // Add right border to serial cell
                                    serialCell.borderBottom = '1px solid black'; // Add bottom border
                                    serialCell.borderTop = '1px solid black'; // Add top border
                                    serialCell.borderLeft = '3px solid #f1c40f'; // Add YELLOW left border
                                } else {
                                    redCount++;
                                    redSerials.push(serialId);
                                    serialCell.borderRight = '3px solid  #e74c3c'; // Add right border to serial cell
                                    serialCell.borderBottom = '1px solid black'; // Add bottom border
                                    serialCell.borderTop = '1px solid black'; // Add top border
                                    serialCell.borderLeft = '3px solid  #e74c3c'; // Add RED left border
                                }

                                // Update the color-coded boxes and counts in the container
                                updateSanitizationCounts(greenCount, yellowCount, redCount, rows.length, greenSerials, yellowSerials, redSerials);
                            } else {
                                console.error('SanReq Error: Sanitization requirement element not found in HTML response');
                            }
                        } else {
                            console.error('SanReq Error: GM_xmlhttpRequest failed with status', response.status);
                        }
                    },
                    onerror: function (error) {
                        console.error('SanReq Error: GM_xmlhttpRequest failed with error', error);
                    }
                });

            });
        } catch (error) {
            console.error('SanReq Error fetching or processing data:', error);
        }
    };

    // Function to update the color-coded boxes and counts in the container
    const updateSanitizationCounts = (greenCount, yellowCount, redCount, totalRows, greenSerials, yellowSerials, redSerials) => {
        // Function to generate the URL based on serials
        const generateURL = (serials) => {
            const baseUrl = 'https://mobility.AMAZON.com/part/search?search_type=all&search_string=serialId&max_rows=50&query=GO';
            const serialsParam = serials.join('+');
            return baseUrl.replace('serialId', serialsParam);
        };

        // Create the color-coded boxes and counts with 3D/shadow effect
        const colorBoxes = `
    <style>
        .color-box-container {
            display: flex;
            align-items: center;
            justify-content: center;
            margin-bottom: 10px;
        }

        .color-box {
            display: flex;
            align-items: center;
            justify-content: center;
            width: 60px; /* Adjusted width */
            height: 30px;
            border-radius: 8px;
            font-size: 12px;
            padding: 4px;
            margin-right: 1px; /* Reduced margin between color boxes */
            margin-top: 1px;
            margin-bottom: 3px;
            border: 4px solid #ccc; /* Lighter border color */
            border-top: 4px solid #f5f5f5; /* Lighter color on top for shine effect */
            border-left: 4px solid #f5f5f5; /* Lighter color on left for shine effect */
            background-clip: padding-box; /* Apply background within padding */
            font-family: Arial; /* Set font family with fallbacks */
            font-weight: bold; /* Bold text */
            text-shadow: 1px 1px 2px rgba(255, 255, 255, 0.5); /* Soft text shadow for 3D effect */
        }

        .green-box {
            background-color:  #2ecc71;
            color: BLACK;
        }

        .yellow-box {
            background-color: #f1c40f;
            color: black;
        }

        .red-box {
            background-color: #e74c3c;
            color: BLACK;
        }

        /* Notification tag styles */
        .notification-tag {
            position: fixed;
            bottom: -30px;
            left: 50%;
            font-size: 16px;
            font-weight: bold;
            transform: translateX(-50%);
            background-color: #3498db;
            color: black;
            padding: 8px 16px;
            border-radius: 20px;
            transition: bottom 0.3s ease-in-out;
            z-index: 1000;
            border: 2px solid black; /* Added black border */
        }

        .show-notification-tag {
            bottom: 10px;
        }

        .green-notification {
            background-color: #2ecc71;
        }

        .yellow-notification {
            background-color: #f4d03f;
        }

        .red-notification {
            background-color: #ec7063;
        }
    </style>
   <div class="color-box-container">
        ${greenCount > 0 ? `
            <div class="color-box green-box clickable-color-box" data-serials="${greenSerials.join(',')}" title="🔗Open Mobility link in new tab for these ${greenCount} No San Required serials">
                No San<br>${greenCount} / ${totalRows}
            </div>
        ` : ''}
        ${yellowCount > 0 ? `
            <div class="color-box yellow-box clickable-color-box" data-serials="${yellowSerials.join(',')}" title="🔗Open Mobility link in new tab for these ${yellowCount} San Required serials">
                San Req<br>${yellowCount} / ${totalRows}
            </div>
        ` : ''}
        ${redCount > 0 ? `
            <div class="color-box red-box clickable-color-box" data-serials="${redSerials.join(',')}" title="🔗Open Mobility link in new tab for these ${redCount} Unknown serials">
                Unknown<br>${redCount} / ${totalRows}
            </div>
        ` : ''}
    </div>
    `;

        // Check if all rows have been processed
        if (greenCount + yellowCount + redCount === totalRows) {
            let bannerText = `Mixed San Status (Qty: ${totalRows})⚠️`;
            const notificationTag = document.createElement('div');
            notificationTag.classList.add('notification-tag', 'show-notification-tag');

            if (greenCount > 0 && yellowCount === 0 && redCount === 0) {
                bannerText = `No San Required for all ${totalRows} Displayed Serials ✔️`;
                notificationTag.classList.add('green-notification');
            }

            if (greenCount === 0 && yellowCount > 0 && redCount === 0) {
                bannerText = `San Required for all ${totalRows} Displayed Serials ⚠️`;
                notificationTag.classList.add('yellow-notification');
            }

            if (greenCount === 0 && yellowCount === 0 && redCount > 0) {
                bannerText = `Unknown San Status for all ${totalRows} Displayed Serials ❗`;
                notificationTag.classList.add('red-notification');
            }

            notificationTag.innerText = bannerText;
            document.body.appendChild(notificationTag);

            // Hide the notification tag after 3 seconds
            setTimeout(() => {
                notificationTag.classList.remove('show-notification-tag');
                setTimeout(() => {
                    notificationTag.remove();
                }, 300); // Remove the notification tag after transition ends
            }, 6000); // Show the notification for 3 seconds
        }

        // Check if the color boxes already exist in the navbar inner container
        const existingColorBoxes = document.querySelector('.navbar-inner .color-box-container');
        if (existingColorBoxes) {
            // Color boxes already exist, update the counts
            existingColorBoxes.innerHTML = colorBoxes;
        } else {
            // Color boxes don't exist, create and append them to the navbar inner container
            const navbarInner = document.querySelector('.navbar-inner');
            if (!navbarInner) {
                console.error('Error: Navbar inner container not found');
                return;
            }

            // Append the color boxes to the navbar inner container
            navbarInner.innerHTML += colorBoxes;

            // Event delegation to handle color box click events
            navbarInner.addEventListener('click', (event) => {
                const target = event.target;
                if (target.classList.contains('clickable-color-box')) {
                    const serials = target.dataset.serials.split(',');
                    const url = generateURL(serials);
                    window.open(url, '_blank');
                }
            });
        }
    };

    // Example usage: Call the fetchSanRequirementStepByStep function
    fetchSanRequirementStepByStep();

    // Function to create the container for sanitization counts
    const createSanitizationCountsContainer = () => {
        // Create the container element
        const container = document.createElement('div');
        container.id = 'sanitizationCountsContainer'; // Set the ID for the container

        // Append the container to the body of the document
        document.body.appendChild(container);
    };

    // Example usage: Call the function to create the container
    createSanitizationCountsContainer();









    ////////////▌│█║▌║▌║ SCAN SN INPUT ▌│█║▌║▌║//////////////////////


    // Create input field, button, and tally board
    const scannerInput = document.createElement('input');
    scannerInput.type = 'text';
    scannerInput.placeholder = 'SCAN MULTIPLE ▌│█║▌║▌║';
    scannerInput.style.width = '200px';
    scannerInput.style.height = '30px';
    scannerInput.style.borderRadius = '5px';
    scannerInput.style.marginLeft = '15px';
    scannerInput.style.fontSize = '16px';
    scannerInput.style.marginTop = '8px';
    scannerInput.title = '⚠️ Lets you scan multiple SNs into a table that you can then bulk search in Mobility';

    const searchButton = document.createElement('button');
    searchButton.textContent = '🔎';
    searchButton.style.height = '30px';
    searchButton.style.fontSize = '16px';
    searchButton.style.marginBottom = '1px';
    searchButton.style.backgroundColor = 'transparent'; // Remove background color
    searchButton.style.border = 'none'; // Remove border


    const scannedList = document.createElement('div');
    scannedList.style.marginBottom = '20px';

    let scannedSerialNumbers = [];

    function createDeleteButton(scannedSN) {
        const deleteButton = document.createElement('button');
        deleteButton.textContent = '❌';
        deleteButton.style.color = 'red';
        deleteButton.style.backgroundColor = 'transparent';
        deleteButton.style.border = 'none';
        deleteButton.style.cursor = 'pointer';
        deleteButton.style.fontSize = '14px';
        deleteButton.style.marginLeft = '1px';

        deleteButton.addEventListener('mouseenter', function() {
            deleteButton.style.transform = 'scale(1.1)';
            deleteButton.style.transition = 'transform 0.2s ease';
        });

        deleteButton.addEventListener('mouseleave', function() {
            deleteButton.style.transform = 'scale(1)';
        });

        deleteButton.addEventListener('click', function() {
            deleteScannedSerialNumber(scannedSN);
        });

        return deleteButton;
    }





    /// FETCHING SERIAL MODELS //////



    let modelList = []; // Define modelList as a global variable
    let modelMap = new Map(); // Keep track of fetched models
    let modelValueLogged = false; // Define modelValueLogged as a global variable



    // Function to update model cell content and style
    function updateModelCell(modelCell, value) {
        modelCell.textContent = value;
        modelCell.style.backgroundColor = 'WHITE'; // White background color for default models
    }

    // Function to check if the model value is valid
    function isValidModel(modelValue) {
        // Check if the modelValue matches the pattern of "Number of Search Results"
        return !modelValue.match(/^Number of Search Results:/);
    }

    // Define the processScannedSerial function
    function processScannedSerial(serial) {
        console.log('Serial scanned:', serial);
    }

    // Simulated scanning of serials
    setTimeout(() => {
        fetchModelData('Serial1', document.getElementById('modelCell1'));
    }, 1000); // Simulate scanning after 1 second
    setTimeout(() => {
        fetchModelData('Serial2', document.getElementById('modelCell2'));
    }, 2000); // Simulate scanning after 2 seconds





    function updateScannedList() {
        scannedList.innerHTML = '';

        // Create a heading container for serials and models
        const headingsContainer = document.createElement('div');
        headingsContainer.style.display = 'flex';
        headingsContainer.style.position = 'sticky'; // Fix position when scrolling
        headingsContainer.style.top = '0'; // Stick to top
        headingsContainer.style.backgroundColor = 'rgba(35, 47, 63, 0.9)'; // Match board background
        headingsContainer.style.padding = '5px 0'; // Add some padding
        headingsContainer.style.zIndex = '100'; // Ensure it stays above other elements

        // Create a heading for serial cells
        const serialsHeading = document.createElement('div');
        serialsHeading.textContent = 'Serials';
        serialsHeading.style.fontWeight = 'bold';
        serialsHeading.style.fontSize = '14px';
        serialsHeading.style.marginBottom = '-10px';
        headingsContainer.appendChild(serialsHeading);

        // Create a heading for model cells
        const modelHeading = document.createElement('div');
        modelHeading.textContent = 'Model';
        modelHeading.style.fontWeight = 'bold';
        modelHeading.style.fontSize = '14px';
        modelHeading.style.marginLeft = '150px';
        modelHeading.style.marginBottom = '5px';
        headingsContainer.appendChild(modelHeading);

        scannedList.appendChild(headingsContainer);

        scannedSerialNumbers.forEach(function (scannedSN, index) {
            const container = document.createElement('div');
            container.style.display = 'flex';
            container.style.alignItems = 'center';

            const bubble = document.createElement('a');
            bubble.href = 'https://mobility.AMAZON.com/part/search?search_type=all&search_string=' + encodeURIComponent(scannedSN) + '&max_rows=50&query=GO';
            bubble.target = '_blank';
            bubble.textContent = `${index + 1}. ${scannedSN}`;
            bubble.classList.add('bubble');
            bubble.style.marginBottom = '5px';

            // Default to blue for all serials
            bubble.style.backgroundColor = '#3498db';

            const deleteButton = createDeleteButton(scannedSN);

            const modelCell = document.createElement('div');
            modelCell.style.marginLeft = '10px';
            modelCell.style.color = 'black';
            modelCell.style.backgroundColor = 'white'; // Corrected backgroundColor property name
            modelCell.style.fontSize = '14px';
            modelCell.style.fontWeight = 'bold';
            modelCell.style.fontFamily = 'Arial';
            modelCell.style.padding = '5px 10px';
            modelCell.style.borderRadius = '5px';
            modelCell.style.marginBottom = '5px';
            modelCell.style.marginLeft = '-2px';
            modelCell.style.border = '1px solid black';

            fetchModelData(scannedSN, modelCell);

            const containerWrapper = document.createElement('div');
            containerWrapper.style.display = 'flex';
            containerWrapper.style.alignItems = 'center';
            containerWrapper.appendChild(bubble);
            containerWrapper.appendChild(modelCell);

            container.appendChild(containerWrapper);
            container.appendChild(deleteButton);

            scannedList.appendChild(container);

            // Check if the serial matches any in the table
            const tableRows = document.querySelectorAll('.table-bordered.table-striped.table-condensed tbody tr');
            tableRows.forEach(row => {
                let serialCell = row.cells[2];
                if (serialCell) {
                    let serialText = serialCell.textContent.trim().toUpperCase();
                    serialText = serialText.replace(/\s*\(\s*RMA\s*\)\s*$/, ''); // Properly remove "( RMA )" text

                    if (serialText === scannedSN) {
                        bubble.style.backgroundColor = 'lightgreen'; // Match found, set to lightgreen
                        serialCell.style.backgroundColor = 'lightgreen'; // Highlight matched serial in the table
                        row.style.backgroundColor = 'lightgreen'; // Highlight the entire row
                    }
                }
            });
        });

        const style = document.createElement('style');
        style.textContent = `
.bubble {
    font-size: 14px;
    font-family: Arial;
    font-weight: bold;
    padding: 5px 10px;
    border-radius: 5px;
    background-color: #3498db;
    color: black;
    text-decoration: none;
    cursor: pointer;
    border: 1px solid black; /* Add black border */
}

.bubble:hover {
    background-color: #FF7F7F;
}
`;
        document.head.appendChild(style);

        let quantity = scannedSerialNumbers.length;
        title.textContent = `Scanned Qty = ${quantity}`;
        const bulkDeleteButton = document.createElement('button');
        bulkDeleteButton.textContent = '❌';
        bulkDeleteButton.style.marginLeft = '200px';
        bulkDeleteButton.style.fontSize = '20px';
        bulkDeleteButton.style.cursor = 'pointer';
        bulkDeleteButton.style.backgroundColor = 'transparent'; // Remove background color
        bulkDeleteButton.style.border = 'none'; // Remove border

        bulkDeleteButton.addEventListener('click', function () {
            scannedSerialNumbers = [];
            updateScannedList();
        });

        title.appendChild(searchButton);
        title.appendChild(bulkDeleteButton);

        if (scannedSerialNumbers.length === 0) {
            board.style.display = 'none';
        } else {
            board.style.display = 'block';
        }
    }

    // Function to flash the latest scanned model on screen (only one message at a time)
    function flashLatestScannedModel(model) {
        // Remove any existing flash message before showing a new one
        const existingFlashBanner = document.getElementById('flashBanner');
        if (existingFlashBanner) {
            document.body.removeChild(existingFlashBanner); // Remove the old flash message
        }

        // Create the new flash banner
        const flashBanner = document.createElement('div');
        flashBanner.id = 'flashBanner'; // Set an ID to manage it later
        flashBanner.innerHTML = `<strong>${model}</strong>`; // Make the model text bold using <strong>
        flashBanner.style.position = 'fixed';
        flashBanner.style.top = '20px';
        flashBanner.style.left = '50%';
        flashBanner.style.transform = 'translateX(-50%)';
        flashBanner.style.backgroundColor = '#aed6f1';
        flashBanner.style.color = 'black';
        flashBanner.style.fontSize = '20px';
        flashBanner.style.padding = '10px';
        flashBanner.style.borderRadius = '5px';
        flashBanner.style.zIndex = '10000';
        flashBanner.style.boxShadow = '0 4px 8px rgba(0,0,0,0.4)';

        // Append the banner to the document body
        document.body.appendChild(flashBanner);


        // Append the banner to the document body
        document.body.appendChild(flashBanner);

        // Remove the flash message after 2 seconds
        setTimeout(() => {
            if (flashBanner) {
                document.body.removeChild(flashBanner);
            }
        }, 2000); // Display for 2 seconds
    }

    // Update fetchModelData to flash the latest scanned model and leave older models untouched
    function fetchModelData(serialId, modelCell) {
        if (!modelCell) {
            console.error('Error: modelCell is undefined');
            return;
        }

        // Check if the model has already been fetched and cached
        if (modelMap.has(serialId)) {
            const cachedModel = modelMap.get(serialId).modelValue;
            updateModelCell(modelCell, cachedModel);
            return;
        }

        const xhr = new XMLHttpRequest();
        const url = `https://mobility.AMAZON.com/part/search?search_type=all&search_string=${serialId}&max_rows=6000&query=GO`;

        xhr.onreadystatechange = function() {
            if (xhr.readyState === XMLHttpRequest.DONE) {
                if (xhr.status === 200) {
                    const response = xhr.responseText;
                    console.log('HTML Response:', response); // Log HTML response for debugging

                    // Check if response contains "no results" string
                    if (response.includes('no results')) {
                        modelCell.textContent = 'NOT FOUND';
                        modelCell.style.backgroundColor = '#FF5733'; // Red background color
                    } else {
                        const parser = new DOMParser();
                        const xmlDoc = parser.parseFromString(response, 'text/html');
                        const listItems = xmlDoc.querySelectorAll('li');
                        let modelValue = ''; // Initialize modelValue

                        // Find the model value
                        listItems.forEach((item) => {
                            if (item.textContent.includes('Model') && !item.textContent.includes('Model Apn')) {
                                const nextItem = item.nextElementSibling;
                                if (nextItem) {
                                    modelValue = nextItem.textContent.trim().replace(/\s+\(\d+\)$/, ''); // Remove " (1)" at the end
                                    updateModelCell(modelCell, modelValue); // Update model cell with the model value
                                    modelMap.set(serialId, { modelValue }); // Cache the model data for future use
                                    modelCell.classList.add('processed'); // Mark the cell as processed

                                    flashLatestScannedModel(modelValue); // Flash the latest scanned model on screen

                                    return; // Exit the loop once the model value is found
                                }
                            }
                        });

                        if (!isValidModel(modelValue)) {
                            modelCell.textContent = 'NOT FOUND';
                            modelCell.style.backgroundColor = '#FF5733'; // Red background color
                        }
                    }
                } else {
                    console.error('Error fetching data:', xhr.status);
                }
            }
        };

        xhr.open('GET', url, true);
        xhr.send();
    }


    function deleteScannedSerialNumber(scannedSN) {
        const index = scannedSerialNumbers.indexOf(scannedSN);
        if (index !== -1) {
            scannedSerialNumbers.splice(index, 1);

            // Remove green highlight from table if the serial was a match
            const tableRows = document.querySelectorAll('.table-bordered.table-striped.table-condensed tbody tr');
            tableRows.forEach(row => {
                let serialCell = row.cells[2];
                if (serialCell) {
                    let serialText = serialCell.textContent.trim().toUpperCase();
                    serialText = serialText.replace(/\s*\(\s*RMA\s*\)\s*$/, ''); // Properly remove "( RMA )" text

                    if (serialText === scannedSN) {
                        serialCell.style.backgroundColor = ''; // Remove the green highlight
                        row.style.backgroundColor = ''; // Remove the row highlight
                    }
                }
            });

            updateScannedList();
        }
    }

    function searchSerialNumbers() {
        if (scannedSerialNumbers.length > 0) {
            const url = 'https://mobility.AMAZON.com/part/search?search_type=all&search_string=' + scannedSerialNumbers.join('+') + '&max_rows=6000&query=GO';
            window.open(url, '_blank');
        } else {
            alert('No serial numbers scanned!');
        }
    }

    function showRedBanner2(message) {
        const existingBanner = document.getElementById('errorBanner');
        if (existingBanner) {
            existingBanner.textContent = message;
        } else {
            const banner = document.createElement('div');
            banner.id = 'errorBanner';
            banner.textContent = message;
            banner.style.position = 'fixed';
            banner.style.top = '0';
            banner.style.left = '0';
            banner.style.width = '100%';
            banner.style.fontSize = '20px';
            banner.style.backgroundColor = '#ff4d4d';
            banner.style.color = 'white';
            banner.style.textAlign = 'center';
            banner.style.padding = '10px';
            banner.style.zIndex = '9999';
            document.body.appendChild(banner);
            setTimeout(function () {
                document.body.removeChild(banner);
            }, 1000);
        }
    }

    // Modify scanning behavior to always check for matches
    scannerInput.addEventListener('keydown', function (event) {
        if (event.key === 'Enter') {
            event.preventDefault();
            const scannedSN = scannerInput.value.trim().toUpperCase();
            const hasSpecialCharacters = /[^A-Z0-9\-_]/.test(scannedSN); // Allow alphanumeric characters, hyphens, and underscores
            const isDuplicate = scannedSerialNumbers.includes(scannedSN);

            if (scannedSN && !hasSpecialCharacters) {
                if (!isDuplicate) {
                    scannedSerialNumbers.unshift(scannedSN);

                    // Always check for matches with existing serials in the table
                    console.log("🎯 SERIAL MATCH: Checking for matches...");
                    const tableRows = document.querySelectorAll('.table-bordered.table-striped.table-condensed tbody tr');
                    let serialFound = false;

                    tableRows.forEach(row => {
                        let serialCell = row.cells[2];
                        if (serialCell) {
                            let serialText = serialCell.textContent.trim().toUpperCase();
                            serialText = serialText.replace(/\s*\(\s*RMA\s*\)\s*$/, '');

                            console.log(`🎯 SERIAL MATCH: Comparing scanned serial "${scannedSN}" with table serial "${serialText}"`);

                            if (serialText === scannedSN) {
                                serialFound = true;
                                serialCell.style.backgroundColor = 'lightgreen'; // Highlight matched serial in the table
                                row.style.backgroundColor = 'lightgreen'; // Highlight the entire row
                                console.log(`🎯 SERIAL MATCH: Match found for serial "${scannedSN}"`);
                            }
                        }
                    });

                    updateScannedList(); // Update the scanned list with color coding
                    scannerInput.value = '';
                } else {
                    console.log(`🎯 SERIAL MATCH: Duplicate serial "${scannedSN}" detected`);
                    showRedBanner2('DUPLICATE SCAN ❗');
                    scannerInput.value = '';
                }
            } else {
                console.log(`🎯 SERIAL MATCH: Invalid serial "${scannedSN}" due to special characters or empty input`);
                showRedBanner2('NO SPECIAL CHARACTERS ❗');
                scannerInput.value = '';
            }
        }
    });





    searchButton.addEventListener('click', function () {
        if (scannedSerialNumbers.length > 0) {
            const url = 'https://mobility.AMAZON.com/part/search?search_type=all&search_string=' + scannedSerialNumbers.join('+') + '&max_rows=6000&query=GO';
            window.open(url, '_blank');
        } else {
            alert('No serial numbers scanned!');
        }
    });

    const board = document.createElement('div');
    board.style.position = 'fixed';
    board.style.top = '50px';
    board.style.left = '50px';
    board.style.width = '450px';
    board.style.color = 'white';
    board.style.padding = '10px';
    board.style.backgroundColor = 'rgba(35, 47, 63, 0.8)';
    board.style.borderRadius = '10px';
    board.style.boxShadow = '0 4px 8px rgba(0,0,0,0.4)';
    board.style.border = '1px solid rgba(255, 255, 255, 0.2)';
    board.style.zIndex = '9999';
    board.style.overflowY = 'auto';
    board.style.maxHeight = '450px';
    board.style.display = 'none';

    const title = document.createElement('div');
    title.textContent = 'Scanned Qty = 0';
    title.style.fontSize = '20px';
    title.style.fontWeight = 'bold';
    title.style.marginBottom = '10px';
    board.appendChild(title);

    board.appendChild(scannedList);
    document.body.appendChild(board);











    //// ADD SHIPMENT ID  🚚🔁//////


    // Define the input textbox globally
    const shipmentIdinputTextbox = document.createElement('input');

    function shipmentIdInput() {
        shipmentIdinputTextbox.type = 'text';
        shipmentIdinputTextbox.placeholder = 'TRACKING#🚚';
        shipmentIdinputTextbox.style.textAlign = 'center'; // Center text




        // Add hover description
        shipmentIdinputTextbox.title = '⚠️ Adds Tracking ID to checked ☑️ SNs';

        // Function to show the yellow banner with moving truck emoji for 3 seconds
        const shipmentChangeSuccessBanner = (shipmentChangeSuccessMessage) => {
            console.log('Show success banner:', shipmentChangeSuccessMessage);

            // Create the yellow banner element
            const shipmentIdbanner = document.createElement('div');
            shipmentIdbanner.innerHTML = shipmentChangeSuccessMessage;
            shipmentIdbanner.style.position = 'fixed';
            shipmentIdbanner.style.top = '0';
            shipmentIdbanner.style.left = '0';
            shipmentIdbanner.style.width = '100%';
            shipmentIdbanner.style.fontSize = '20px';
            shipmentIdbanner.style.backgroundColor = 'yellow';
            shipmentIdbanner.style.color = 'black';
            shipmentIdbanner.style.textAlign = 'center';
            shipmentIdbanner.style.padding = '10px';
            shipmentIdbanner.style.zIndex = '9999';
            shipmentIdbanner.style.border = '1px solid black'; // Add black border

            // Append the banner to the document body
            document.body.appendChild(shipmentIdbanner);

            // After 3 seconds, remove the banner
            setTimeout(() => {
                document.body.removeChild(shipmentIdbanner);
            }, 3000);
        };

        // Function to show the red error banner for 6 seconds
        const shipmentChangeErrorBanner = (errorMessage) => {
            console.log('Show error banner:', errorMessage);

            // Create the red banner element
            const shipmentIdErrorbanner = document.createElement('div');
            shipmentIdErrorbanner.innerHTML = errorMessage;
            shipmentIdErrorbanner.style.position = 'fixed';
            shipmentIdErrorbanner.style.top = '0';
            shipmentIdErrorbanner.style.left = '0';
            shipmentIdErrorbanner.style.width = '100%';
            shipmentIdErrorbanner.style.fontSize = '20px';
            shipmentIdErrorbanner.style.backgroundColor = 'red';
            shipmentIdErrorbanner.style.color = 'white';
            shipmentIdErrorbanner.style.textAlign = 'center';
            shipmentIdErrorbanner.style.padding = '10px';
            shipmentIdErrorbanner.style.zIndex = '9999';
            shipmentIdErrorbanner.style.border = '1px solid black'; // Add black border

            // Append the banner to the document body
            document.body.appendChild(shipmentIdErrorbanner);

            // After 6 seconds, remove the banner
            setTimeout(() => {
                document.body.removeChild(shipmentIdErrorbanner);
            }, 3000);
        };

        function changeTrackingId() {
            shipmentIdinputTextbox.addEventListener('keydown', function (event) {
                if (event.keyCode === 13) { // Check if Enter key is pressed
                    const trackingId = shipmentIdinputTextbox.value.trim();
                    console.log('Entered tracking ID:', trackingId);

                    // Check if the tracking ID input is empty
                    if (trackingId === '') {
                        // Display error message if tracking ID is empty
                        console.log('⚠️ Please enter Tracking ID!');
                        shipmentChangeErrorBanner('⚠️ Please enter <strong>Tracking ID</strong>!');
                        return;
                    }

                    // Find all checkbox elements by searching for their title attribute
                    const checkboxElements = document.querySelectorAll('input[type="checkbox"][title="Select checkbox to mark this part for Bulk Edit"]');
                    const checkboxesChecked = Array.from(checkboxElements).filter(checkbox => checkbox.checked);
                    let checkboxesToChange = checkboxElements.length;

                    // Check if any checkbox is checked
                    if (checkboxElements.length === 0 || !Array.from(checkboxElements).some(checkbox => checkbox.checked)) {
                        shipmentChangeErrorBanner('⚠️ Please select☑️ parts to change!');
                        return;
                    }

                    // Log the number of selected rows
                    console.log('Number of parts selected to change:', checkboxesToChange);

                    // Select tracking ID checkbox
                    const trackingIdCheckbox = document.getElementById('edit_tracking_id_checkbox');
                    if (trackingIdCheckbox) {
                        trackingIdCheckbox.checked = true;
                        console.log('Tracking ID checkbox checked');

                        // Append text to the tracking ID textbox
                        const trackingIdTextbox = document.getElementById('tracking_id_textbox');
                        if (trackingIdTextbox) {
                            // Append the entered tracking ID to the tracking ID textbox
                            trackingIdTextbox.value = trackingId;
                            console.log('Tracking ID Textbox value set:', trackingId);

                            // Trigger the "Make Changes" button using XPath
                            const container = document.getElementById('part_search_bulk_edit_popup_dialog_form');
                            if (container) {
                                const makeChangesButtonXPath = `//div[@class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix"]//button[contains(text(), "Make Changes")]`;
                                const makeChangesButton = document.evaluate(makeChangesButtonXPath, container, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

                                if (makeChangesButton) {
                                    makeChangesButton.click();
                                    console.log('Make Changes button clicked');


                                    // Function to periodically check the background color of all checkbox elements - SHOW BANNER WHEN ALL CELLS GREEN
                                    const checkBackgroundColorTrackingId = setInterval(() => {
                                        let greenCheckboxes = 0;
                                        checkboxElements.forEach((checkboxElement) => {
                                            // Get the parent td element
                                            const tdElement = checkboxElement.closest('td');
                                            // Log the background color of the td element
                                            if (tdElement && window.getComputedStyle(tdElement).getPropertyValue('background-color') === 'rgb(50, 205, 50)') {
                                                greenCheckboxes++;
                                            }
                                        });

                                        if (greenCheckboxes === checkboxesToChange) {
                                            // All checkboxes have turned green, display success message
                                            shipmentChangeSuccessBanner(`x<strong>${checkboxesToChange}</strong> part(s) Tracking ID changed to: <strong>${trackingId}</strong>`);
                                            clearInterval(checkBackgroundColorTrackingId);

                                            // Reset cell colors
                                            checkboxElements.forEach((checkboxElement) => {
                                                const tdElement = checkboxElement.closest('td');
                                                if (tdElement) {
                                                    tdElement.style.backgroundColor = 'inherit';
                                                }
                                            });


                                            // Uncheck tracking ID checkbox
                                            trackingIdCheckbox.checked = false;
                                            console.log('Tracking ID checkbox unchecked');

                                            // Clear tracking ID textbox
                                            trackingIdTextbox.value = '';
                                            console.log('Tracking ID Textbox cleared');
                                            // Uncheck the "Bulk Edit Select All" checkbox if it was selected before running the function
                                            const bulkEditSelectAllCheckbox = document.getElementById('part_search_bulk_edit_select_all');
                                            if (bulkEditSelectAllCheckbox && bulkEditSelectAllCheckbox.checked) {
                                                bulkEditSelectAllCheckbox.checked = false;
                                                console.log('Bulk edit select all checkbox unchecked');
                                            }
                                        }
                                    }, 500); // Check every 0.5 seconds
                                } else {
                                    console.log('Make Changes button not found!');
                                }
                            } else {
                                console.log('Container element not found!');
                            }
                        } else {
                            console.log('Outbound RMA ID Textbox not found!');
                        }
                    } else {
                        console.log('Outbound RMA ID checkbox not found!');
                    }
                    // Clear the input textbox
                    shipmentIdinputTextbox.value = '';
                }
            });
        }



        // Call the function to listen for changes and "Enter" key press
        changeTrackingId();


    }

    // Call the function to create the input textbox and listen for changes and "Enter" key press
    shipmentIdInput();

    // Append the input textbox to the document body
    document.body.appendChild(shipmentIdinputTextbox);
















    ////🗑️🔁 CHANGE BIN NAME /////


    // Define the input textbox globally
    const BinInputTextbox = document.createElement('input');

    function BinIdInput() {
        BinInputTextbox.type = 'text';
        BinInputTextbox.placeholder = 'BIN 🗑️';
        BinInputTextbox.style.textAlign = 'center'; // Center text


        // Add hover description
        BinInputTextbox.title = '⚠️ Transfers checked ☑️ SNs to BIN';

        // Function to show the yellow banner with moving truck emoji for 3 seconds
        const BinSuccessBanner = (BinMessage) => {
            console.log('Show success banner:', BinMessage);

            // Create the BIN Success banner element
            const BinBanner = document.createElement('div');
            BinBanner.innerHTML = BinMessage;
            BinBanner.style.position = 'fixed';
            BinBanner.style.top = '0';
            BinBanner.style.left = '0';
            BinBanner.style.width = '100%';
            BinBanner.style.fontSize = '20px';
            BinBanner.style.backgroundColor = '#e5e7e9';
            BinBanner.style.color = 'black';
            BinBanner.style.textAlign = 'center';
            BinBanner.style.padding = '10px';
            BinBanner.style.zIndex = '9999';
            BinBanner.style.border = '1px solid black'; // Add black border

            // Append the banner to the document body
            document.body.appendChild(BinBanner);

            // After 3 seconds, remove the banner
            setTimeout(() => {
                document.body.removeChild(BinBanner);
            }, 4000);
        };

        // Function to show the red error banner for 6 seconds
        const BinErrorBanner = (errorMessage) => {
            console.log('Show error banner:', errorMessage);

            // Create the red banner element
            const BinErrorBanner = document.createElement('div');
            BinErrorBanner.innerHTML = errorMessage;
            BinErrorBanner.style.position = 'fixed';
            BinErrorBanner.style.top = '0';
            BinErrorBanner.style.left = '0';
            BinErrorBanner.style.width = '100%';
            BinErrorBanner.style.fontSize = '20px';
            BinErrorBanner.style.backgroundColor = 'red';
            BinErrorBanner.style.color = 'white';
            BinErrorBanner.style.textAlign = 'center';
            BinErrorBanner.style.padding = '10px';
            BinErrorBanner.style.zIndex = '9999';
            BinErrorBanner.style.border = '1px solid black'; // Add black border

            // Append the banner to the document body
            document.body.appendChild(BinErrorBanner);

            // After 6 seconds, remove the banner
            setTimeout(() => {
                document.body.removeChild(BinErrorBanner);
            }, 4000);
        };




        ///// CHANGE BIN NAME FUNCTION //////////////  14s to 8s  / 6 click to 3 click


        function changeBinId() {
            BinInputTextbox.addEventListener('keydown', function(event) {
                if (event.keyCode === 13) { // Check if Enter key is pressed
                    const inputPattern = /^[A-Za-z]{3}\d{1,3}\.[\w\d\s\S]+?\.[\w\d\s\S]+$/;
                    const BinId = BinInputTextbox.value.trim();
                    console.log('Entered BIN ID:', BinId);

                    // Validate input format
                    if (!inputPattern.test(BinId)) {
                        // Display error message if input format is incorrect
                        BinErrorBanner('⚠️ Invalid input format! Please enter the BIN ID in the format: SITE00.ROOM.BIN');
                        return;
                    }

                    // Find all checkbox elements by searching for their title attribute
                    const checkboxElements = document.querySelectorAll('input[type="checkbox"][title="Select checkbox to mark this part for Bulk Edit"]');
                    const checkboxesChecked = Array.from(checkboxElements).filter(checkbox => checkbox.checked);
                    const checkboxesToChange = checkboxesChecked.length;

                    if (checkboxesToChange === 0) {
                        // Display error message if no row has been selected
                        console.log('⚠️ Please select parts to change!');
                        BinErrorBanner('⚠️ Please select☑️ parts to change!');
                        return;
                    }

                    if (checkboxesChecked.length === 0) {
                        console.log('No rows selected.');
                        return;
                    }

                    // Log the number of selected rows
                    console.log('Number of rows to change:', checkboxesChecked.length);

                    // Select BIN ID checkbox
                    const BinIdCheckbox = document.getElementById('edit_bin_checkbox');
                    if (BinIdCheckbox) {
                        BinIdCheckbox.checked = true;
                        console.log('BIN ID checkbox checked');

                        // Append text to the BIN ID textbox
                        const BinIdTextbox = document.getElementById('bin_asset_tag_textbox');
                        if (BinIdTextbox) {
                            // Append the entered BIN ID to the BIN ID textbox
                            BinIdTextbox.value = BinId;
                            console.log('BIN ID Textbox value set:', BinId);

                            // Trigger the "Make Changes" button using XPath
                            const container = document.getElementById('part_search_bulk_edit_popup_dialog_form');
                            if (container) {
                                const makeChangesButtonXPath = `//div[@class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix"]//button[contains(text(), "Make Changes")]`;
                                const makeChangesButton = document.evaluate(makeChangesButtonXPath, container, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

                                if (makeChangesButton) {
                                    makeChangesButton.click();
                                    console.log('Make Changes button clicked');

                                    // Function to periodically check the background color of selected checkbox elements
                                    const checkBackgroundColorBin = setInterval(() => {
                                        let redCheckboxes = 0;
                                        let greenCheckboxes = 0;
                                        checkboxElements.forEach((checkboxElement) => {
                                            // Get the parent td element
                                            const tdElement = checkboxElement.closest('td');
                                            // Log the background color of the td element
                                            if (tdElement && window.getComputedStyle(tdElement).getPropertyValue('background-color') === 'rgb(50, 205, 50)') {
                                                greenCheckboxes++;
                                            } else if (tdElement && window.getComputedStyle(tdElement).getPropertyValue('background-color') === 'rgb(255, 0, 0)') {
                                                redCheckboxes++;
                                            }
                                        });

                                        if (greenCheckboxes === checkboxesChecked.length) {
                                            // All checkboxes have turned green, display success message
                                            BinSuccessBanner(`x<strong>${checkboxesChecked.length}</strong> part(s) transferred to: <strong>${BinId}</strong> 🗑️`);
                                            clearInterval(checkBackgroundColorBin);

                                            // Reset cell colors
                                            checkboxElements.forEach((checkboxElement) => {
                                                const tdElement = checkboxElement.closest('td');
                                                if (tdElement) {
                                                    tdElement.style.backgroundColor = 'inherit';
                                                }
                                            });

                                            // Uncheck all selected checkboxes
                                            checkboxesChecked.forEach((checkboxElement) => {
                                                checkboxElement.checked = false;
                                            });

                                            // Reset everything in the form
                                            BinIdCheckbox.checked = false;
                                            BinIdTextbox.value = '';
                                            // Clear the input textbox
                                            BinInputTextbox.value = '';

                                            // Uncheck the "Bulk Edit Select All" checkbox if it was selected before running the function
                                            const bulkEditSelectAllCheckbox = document.getElementById('part_search_bulk_edit_select_all');
                                            if (bulkEditSelectAllCheckbox && bulkEditSelectAllCheckbox.checked) {
                                                bulkEditSelectAllCheckbox.checked = false;
                                                console.log('Bulk edit select all checkbox unchecked');
                                            }
                                        } else if (redCheckboxes > 0 && redCheckboxes + greenCheckboxes === checkboxesChecked.length) {
                                            // Display error message if any cell is red and all cells have changed color
                                            BinErrorBanner(`⚠️ Error: Check Bin Name!`);
                                            clearInterval(checkBackgroundColorBin);


                                        }
                                    }, 500); // Check every 0.5 seconds
                                } else {
                                    console.log('Make Changes button not found!');
                                }
                            } else {
                                console.log('Container element not found!');
                            }
                        } else {
                            console.log('Bin ID Textbox not found!');
                        }
                    } else {
                        console.log('Bin ID checkbox not found!');
                    }
                }
            });
        }

        // Call the function to listen for changes and "Enter" key press
        changeBinId();





    }

    // Call the function to create the input textbox and listen for changes and "Enter" key press
    BinIdInput();

    // Append the input textbox to the document body
    document.body.appendChild(BinInputTextbox);














    ///////////// ☑️ CREATE ESV BUTTON///////////////////////// *Time/Step reduction:


    function openEsvURL() {
        window.open('https://issues.amazon.com/issues/create?template=a0a1fc69-cbd7-4c33-b4c3-74c1b31bed78', '_blank');
    }

    const esvButton = document.createElement('button');
    esvButton.textContent = 'ESV ☑️';
    esvButton.onclick = openEsvURL;
    esvButton.classList.add('esvButton');
    esvButton.setAttribute('title', 'Opens ESV approval template'); // Adding tooltip message




    ////////📥 CI Template BUTTON/////////
    function openCIURL() {
        window.open('https://drive.corp.amazon.com/view/pjbyrne@/AWS_GTPC_CI_PL_Template%20_WW(1).xlsx', '_blank');
    }

    const DownloadCITemplate = document.createElement('button');
    DownloadCITemplate.textContent = 'CI Template 📥';
    DownloadCITemplate.onclick = openCIURL;
    DownloadCITemplate.classList.add('DownloadCITemplate');
    DownloadCITemplate.setAttribute('title', 'Opens CI Template'); // Adding tooltip message







    //////// ⛔RZ Exit Tool///////////////////////////////////////////////////


    function OpenRZExitTool() {
        window.open('https://app.boost.aws.a2z.com/platform/work-requests/create/arn%3Aaws%3Aboost%3A012804863341%3Aworkrequest-template%2FDataCenter%2FRed-Zone-Exit%2F1.0/workflow?priority=3&workflowProperties=%7B%7D&autoSubmit=false', '_blank');
    }

    const RZExitTool = document.createElement('button');
    RZExitTool.textContent = 'RZ EXIT ⛔';
    RZExitTool.onclick = OpenRZExitTool;
    RZExitTool.classList.add('RZExitTool');
    RZExitTool.setAttribute('title', 'Opens RZ Exit Tool (Boost)'); // Adding tooltip message







    /////////  CREATE SAN APPROVAL TEMPLATE + COPY TITLE DATA//////////////////////////////////////////////

    function openSanApprovalURL() {
        const targetColumnsIndices = [8, 9, 17, 3]; // Update with the required column indices
        const tableRows = document.querySelectorAll('table.table-bordered.table-striped.table-condensed tbody tr');
        if (!tableRows.length) {
            console.error('Table rows not found');
            return;
        }

        // Function to extract text content while excluding specific HTML elements
        function extractTextContent(element) {
            return element.textContent.trim().replace(/\t|\n/g, '');
        }

        // Get data from specific columns of the first row
        const rowDataArray = targetColumnsIndices.map(index => {
            const cellData = tableRows[0].querySelectorAll('td')[index - 1];
            if (cellData) {
                let cellText = extractTextContent(cellData);

                // Exclude "(RMA)" from the end of rowDataArray[3]
                if (index === 3 && cellText.endsWith('(RMA)')) {
                    cellText = cellText.slice(0, -5); // Remove the last 5 characters
                }

                return cellText;
            } else {
                return '';
            }
        });

        // Get today's date in MM/DD/YY format
        const today = new Date().toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: '2-digit' }).replace(/\//g, '/');

        // Get the text content from the element with ID 'hello-user' and remove "Hello, "
        const aliasText = document.getElementById('hello-user').textContent.trim().replace('Hello, ', '');

        // Combine the data from specific columns into the desired format
        const dataToCopy = `Use Case #1 - Devices currently in the Red Zone that Require Sanitization for RMA - [${rowDataArray[0]}] - Requestor Login: [${aliasText}] - Request Date: [${today}]`;




        // Copy the formatted data to the clipboard
        navigator.clipboard.writeText(dataToCopy)
            .then(() => {
            console.log('Data copied to clipboard:', dataToCopy);

            // Open the URL after copying the data
            window.open('https://approvals.amazon.com/Template/Details/41044', '_blank');
        })
            .catch(err => {
            console.error('Unable to copy to clipboard:', err);
            alert('Failed to copy data to clipboard');
        });
    }

    const sanApprovalButton = document.createElement('button');
    sanApprovalButton.textContent = 'APPROVAL ✅ ';
    sanApprovalButton.onclick = openSanApprovalURL;
    sanApprovalButton.classList.add('exampleButton');
    sanApprovalButton.setAttribute('title', 'Opens San Approval Template and copies title data'); // Adding tooltip message




    //// CSV LIST BUTTONS /////////



    // CREATE SAN PARTS LIST BUTTON
    const createSanPartsListButton = document.createElement('button');
    createSanPartsListButton.textContent = 'PARTS LIST 💾';
    createSanPartsListButton.onclick = copySpecificColumnsToCSV; // Updated function name
    createSanPartsListButton.setAttribute('data-tooltip', '');
    createSanPartsListButton.classList.add('createSanPartsListButton');
    createSanPartsListButton.setAttribute('title', '⚠️ Downloads Andon Override Exception Parts List populated with ALL displayed serials'); // Adding tooltip message




    // CREATE CI DATA DOWNLOAD LIST BUTTON
    const createCIDownloadButton = document.createElement('button');
    createCIDownloadButton.textContent = 'CI DATA 💾';
    createCIDownloadButton.onclick = copySpecificColumnsToCSV3; // Updated function name
    createCIDownloadButton.setAttribute('data-tooltip', '');
    createCIDownloadButton.classList.add('createCIDownloadButton');
    createCIDownloadButton.setAttribute('title', 'Downloads CSV file with relevant CI data');










    //////////////🛑 BUTTON POSITIONING IN CONTAINER////////////////////////////////////




    //Create div elements for each column of buttons and text labels
    const column1 = document.createElement('div');
    const column2 = document.createElement('div');
    const column3 = document.createElement('div');
    const column4 = document.createElement('div');
    const column5 = document.createElement('div');
    const column6 = document.createElement('div');


    function createStyledLabel(text, color) {
        const label = document.createElement('div');
        label.textContent = text;
        label.style.color = color; // Set text color
        label.style.fontSize = '14px';
        label.style.marginBottom = '5px';
        label.style.marginTop = '-15px';
        label.style.backgroundColor = '#444'; // Set background color to darker grey
        label.style.borderRadius = '1px'; // Adjusted border radius for a smoother look
        label.style.padding = '8px 15px'; // Increased padding for better spacing
        label.style.fontFamily = 'Roboto';
        label.style.border = '1px solid black'; // Add a 1px black border
        label.style.boxShadow = '0px 4px 6px rgba(0, 0, 0, 0.1), 0px 1px 3px rgba(0, 0, 0, 0.08)'; // Added box shadow for the glossy effect
        label.style.textShadow = '0.5px 0.5px 2px black'; // Add a black outline around the letters
        return label;
    }

    // Create styled label elements with different colors
    const label1 = createStyledLabel('MOVE ➡️', '#3498db');
    const label2 = createStyledLabel('CHANGE 🔄', ' #ecf0f1 ');
    const label3 = createStyledLabel('SAN 🇽', '#f1c40f');
    const label4 = createStyledLabel('NON-SAN 🆗', '#28b463');
    const label5 = createStyledLabel('GENERAL 🌐', ' #e74c3c ');
    const label6 = createStyledLabel('Other');










    // Apply flexbox styles to arrange buttons and labels in rows
    column1.style.display = 'flex';
    column1.style.flexDirection = 'column'; // Vertical arrangement
    column1.style.alignItems = 'center'; // Align items in the center

    column2.style.display = 'flex';
    column2.style.flexDirection = 'column'; // Vertical arrangement
    column2.style.alignItems = 'center'; // Align items in the center

    column3.style.display = 'flex';
    column3.style.flexDirection = 'column'; // Vertical arrangement
    column3.style.alignItems = 'center'; // Align items in the center

    column4.style.display = 'flex';
    column4.style.flexDirection = 'column'; // Vertical arrangement
    column4.style.alignItems = 'center'; // Align items in the center

    column5.style.display = 'flex';
    column5.style.flexDirection = 'column'; // Vertical arrangement
    column5.style.alignItems = 'center'; // Align items in the center

    column6.style.display = 'flex';
    column6.style.flexDirection = 'column'; // Vertical arrangement
    column6.style.alignItems = 'center'; // Align items in the center


    // Create and configure the dropdown menu
    const stateChangeDropdown = stateChangeMenu();
    const dropdownMenu2 = TargetBinTOA();
    const dropdownMenu3 = SiteTransferBin();







    //////////// APPEND CODE  ////////////////////////////////////



    ////////🔗 APPEND TO CONTAINER /////////////////////


    // Append text labels and buttons to the columns
    column1.appendChild(label1);
    column1.appendChild(BinInputTextbox);
    column1.appendChild(dropdownMenu2);
    column1.appendChild(dropdownMenu3);

    column2.appendChild(label2);
    column2.appendChild(stateChangeDropdown);
    column2.appendChild(GoneInputTextbox);
    column2.appendChild(BuildInputTextbox);
    column2.appendChild(RmaIdInputTextbox);
    column2.appendChild(shipmentIdinputTextbox);

    column3.appendChild(label3);
    column3.appendChild(copyToClipboardButton);
    column3.appendChild(esvButton);
    column3.appendChild(sanApprovalButton);
    column3.appendChild(createSanPartsListButton);


    column4.appendChild(label4);
    column4.appendChild(RZExitTool);


    column5.appendChild(label5);
    column5.appendChild(createCIDownloadButton);
    column5.appendChild(DownloadCITemplate);


    // Apply flexbox styles to arrange columns in a row
    const rowContainer = document.createElement('div');
    rowContainer.style.display = 'flex';
    rowContainer.style.flexDirection = 'row'; // Horizontal arrangement

    // Append columns to the row container
    rowContainer.appendChild(column1);
    rowContainer.appendChild(column2);
    rowContainer.appendChild(column3);
    rowContainer.appendChild(column4);
    rowContainer.appendChild(column5);

    // Append row container to the form container
    formContainer.appendChild(rowContainer);


    let rowDataArray = []; // Declare rowDataArray outside the loop







    ///////////// STYLE SAN BUTTONS //////////////////
    function buttonStylesSan(button) {
        button.style.backgroundColor = '#f1c40f'; // Set background color
        button.style.border = '1px solid black'; // Add black border
        button.style.height = '36px';
        button.style.color = 'black'; // Set font color to white
        button.style.padding = '6px 16px'; // Adjust padding as needed
        button.style.textAlign = 'center'; // Center text horizontally
        button.style.textDecoration = 'none'; // Remove underline
        button.style.display = 'inline-block'; // Set display to inline-block
        button.style.fontSize = '14px'; // Set font size
        button.style.fontWeight = 'bold'; // Set font weight to bold
        button.style.margin = '2px'; // Adjust margin as needed
        button.style.cursor = 'pointer'; // Set cursor to pointer
        button.style.borderRadius = '5px'; // Set border radius
        button.style.transition = 'background-color 0.3s, box-shadow 0.3s'; // Add transition effect for background color and box-shadow

        // Add hover effect
        button.addEventListener('mouseenter', function() {
            button.style.backgroundColor = '#f9e79f';
            button.style.boxShadow = '0px 8px 10px rgba(0, 0, 0, 0.4)'; // Adjust shadow properties for the hover effect
        });

        button.addEventListener('mouseleave', function() {
            button.style.backgroundColor = '#f1c40f';
            button.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Reset shadow properties
        });
    }


    buttonStylesSan(copyToClipboardButton);
    buttonStylesSan(esvButton);
    buttonStylesSan(sanApprovalButton);
    buttonStylesSan(createSanPartsListButton);




    ///////////// STYLE NON-SAN BUTTONS //////////////////
    function buttonStyleNonSan(button) {
        button.style.backgroundColor = ' #28b463 '; // Set background color
        button.style.border = '1px solid black'; // Add black border
        button.style.height = '36px';
        button.style.color = 'black'; // Set font color to white
        button.style.padding = '6px 16px'; // Adjust padding as needed
        button.style.textAlign = 'center'; // Center text horizontally
        button.style.textDecoration = 'none'; // Remove underline
        button.style.display = 'inline-block'; // Set display to inline-block
        button.style.fontSize = '14px'; // Set font size
        button.style.fontWeight = 'bold'; // Set font weight to bold
        button.style.margin = '2px'; // Adjust margin as needed
        button.style.cursor = 'pointer'; // Set cursor to pointer
        button.style.borderRadius = '5px'; // Set border radius
        button.style.transition = 'background-color 0.3s, box-shadow 0.3s'; // Add transition effect for background color and box-shadow

        // Add hover effect
        button.addEventListener('mouseenter', function() {
            button.style.backgroundColor = '#abebc6';
            button.style.boxShadow = '0px 8px 10px rgba(0, 0, 0, 0.4)'; // Adjust shadow properties for the hover effect
        });

        button.addEventListener('mouseleave', function() {
            button.style.backgroundColor = ' #28b463 ';
            button.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Reset shadow properties
        });
    }


    buttonStyleNonSan(RZExitTool);





    ///////////// STYLE FA LAB BUTTONS //////////////////
    function buttonStyleFA(button) {
        button.style.backgroundColor = '#e74c3c'; // Set background color
        button.style.border = '1px solid black'; // Add black border
        button.style.height = '36px';
        button.style.color = 'black'; // Set font color to white
        button.style.padding = '6px 16px'; // Adjust padding as needed
        button.style.textAlign = 'center'; // Center text horizontally
        button.style.textDecoration = 'none'; // Remove underline
        button.style.display = 'inline-block'; // Set display to inline-block
        button.style.fontSize = '14px'; // Set font size
        button.style.fontWeight = 'bold'; // Set font weight to bold
        button.style.margin = '2px'; // Adjust margin as needed
        button.style.cursor = 'pointer'; // Set cursor to pointer
        button.style.borderRadius = '5px'; // Set border radius
        button.style.transition = 'background-color 0.3s, box-shadow 0.3s'; // Add transition effect for background color and box-shadow

        // Add hover effect
        button.addEventListener('mouseenter', function() {
            button.style.backgroundColor = '#f1948a';
            button.style.boxShadow = '0px 8px 10px rgba(0, 0, 0, 0.4)'; // Adjust shadow properties for the hover effect
        });

        button.addEventListener('mouseleave', function() {
            button.style.backgroundColor = '#e74c3c';
            button.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Reset shadow properties
        });
    }

    //// APPLY SAN BUTTON STYLES /////
    buttonStyleFA(createCIDownloadButton);
    buttonStyleFA(DownloadCITemplate);











    /////////////////🔗 APPEND TO TOOLBAR //////////


    // Find the targetDiv element
    const targetDiv = document.querySelector('div.search_pagination');

    //// TARGETDIV STYLES //////
    targetDiv.style.borderRadius = '10px';
    targetDiv.style.marginRight = '40px';
    targetDiv.style.marginLeft = '-8px';
    targetDiv.style.marginTop = '-8px';
    targetDiv.style.width = '120%'; // Adjust width as needed
    targetDiv.style.boxSizing = 'border-box';
    targetDiv.style.textAlign = 'center';
    targetDiv.style.backgroundColor = '#ccd1d1'; // Adjust background color as needed


    // Add border with glossy shine effect using gradients
    targetDiv.style.borderTop = '4px solid #f5f5f5'; // Lighter color on top
    targetDiv.style.borderLeft = '4px solid #f5f5f5'; // Lighter color on left
    targetDiv.style.borderBottom = '4px solid #cccccc'; // Darker color on bottom
    targetDiv.style.borderRight = '4px solid #cccccc'; // Darker color on right




    ///// SEPARATOR ELEMENT ////////////////
    function createEmoji(emoji) {
        const emojiElement = document.createElement('span');
        emojiElement.textContent = emoji;
        emojiElement.style.fontSize = '20px'; // Adjust margin as needed
        emojiElement.style.marginTop = '5px'; // Adjust margin as needed
        emojiElement.style.marginLeft = '8px'; // Adjust margin as needed
        emojiElement.style.marginRight = '8px'; // Adjust margin as needed

        return emojiElement;
    }


    // Append elements to the targetDiv with dividers between them
    targetDiv.insertBefore(moreToolsTab, targetDiv.firstChild);
    targetDiv.insertBefore(createEmoji('  '), targetDiv.firstChild); // n
    targetDiv.insertBefore(scannerInput, targetDiv.firstChild);
    targetDiv.insertBefore(modelChangeDropdown, targetDiv.firstChild);
    targetDiv.insertBefore(transferOptionsDropdown, targetDiv.firstChild);
    targetDiv.insertBefore(shortenButton, targetDiv.firstChild);
    targetDiv.insertBefore(copySerialsButton, targetDiv.firstChild);
    targetDiv.insertBefore(toggleButton, targetDiv.firstChild);
    targetDiv.insertBefore(toggleLastUserButton, targetDiv.firstChild);
    targetDiv.insertBefore(createEmoji('  '), targetDiv.firstChild); // n
    targetDiv.insertBefore(switchContainer, targetDiv.firstChild);
    targetDiv.insertBefore(ColumnButton, targetDiv.firstChild);
    targetDiv.insertBefore(createEmoji('  '), targetDiv.firstChild); // n
    targetDiv.insertBefore(createEmoji('  '), targetDiv.firstChild); // n









    ///////////// STYLE 1st SET BUTTONS //////////////////

    function applyButtonStyles1(button) {
        button.style.backgroundColor = '#31465E'; // Set background color
        button.style.border = '1px solid black'; // Add black border
        button.style.height = '36px';
        button.style.color = 'white'; // Set font color to white
        button.style.padding = '6px 16px'; // Adjust padding as needed
        button.style.textAlign = 'center'; // Center text horizontally
        button.style.textDecoration = 'none'; // Remove underline
        button.style.display = 'inline-block'; // Set display to inline-block
        button.style.fontSize = '14px'; // Set font size
        button.style.fontWeight = 'bold'; // Set font weight to bold
        button.style.margin = '2px'; // Adjust margin as needed
        button.style.cursor = 'pointer'; // Set cursor to pointer
        button.style.borderRadius = '5px'; // Set border radius
        button.style.transition = 'background-color 0.3s, box-shadow 0.3s'; // Add transition effect for background color and box-shadow

        // Add hover effect
        button.addEventListener('mouseenter', function() {
            button.style.backgroundColor = ' #5d6d7e ';
            button.style.boxShadow = '0px 8px 10px rgba(0, 0, 0, 0.4)'; // Adjust shadow properties for the hover effect
        });

        button.addEventListener('mouseleave', function() {
            button.style.backgroundColor = '#31465E';
            button.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Reset shadow properties
        });
    }

    // APPLY STYLES TO MOBILITY PAGE ELEMENTS
    applyButtonStyles1(toggleLastUserButton);
    applyButtonStyles1(toggleButton);
    applyButtonStyles1(copySerialsButton);
    applyButtonStyles1(shortenButton);





    //////////////////➡️ STYLE MOVE BUTTONS ////////////////////

    function moveDropdownStyle(button) {
        button.style.backgroundColor = '#6495ED'; // LIGHT BLUE
        button.style.border = '1px solid black'; // Add black border
        button.style.width = '115px';
        button.style.height = '36px';
        button.style.color = 'black'; // Set font color to white
        button.style.padding = '6px 16px'; // Adjust padding as needed
        button.style.textAlign = 'center'; // Center text horizontally
        button.style.textDecoration = 'none'; // Remove underline
        button.style.display = 'inline-block'; // Set display to inline-block
        button.style.fontSize = '14px'; // Set font size
        button.style.fontWeight = 'bold'; // Set font weight to bold
        button.style.margin = '2px'; // Adjust margin as needed
        button.style.cursor = 'pointer'; // Set cursor to pointer
        button.style.borderRadius = '5px'; // Set border radius
        button.style.transition = 'background-color 0.3s, box-shadow 0.3s'; // Add transition effect for background color and box-shadow
        button.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Add box shadow for 3D effect

        // Add hover effect
        button.addEventListener('mouseenter', function() {
            button.style.backgroundColor = ' #aed6f1 ';
            button.style.boxShadow = '0px 8px 10px rgba(0, 0, 0, 0.4)'; // Adjust shadow properties for the hover effect
        });

        button.addEventListener('mouseleave', function() {
            button.style.backgroundColor = '#6495ED';
            button.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Reset shadow properties
        });
    }


    function moveTextboxStyle(button) {
        button.style.backgroundColor = '#6495ED'; // LIGHT BLUE
        button.style.border = '1px solid black'; // Add black border
        button.style.width = '100px';
        button.style.height = '24px';
        button.style.color = 'black'; // Set font color to white
        button.style.padding = '6px 16px'; // Adjust padding as needed
        button.style.textAlign = 'center'; // Center text horizontally
        button.style.textDecoration = 'none'; // Remove underline
        button.style.display = 'inline-block'; // Set display to inline-block
        button.style.fontSize = '14px'; // Set font size
        button.style.fontWeight = 'bold'; // Set font weight to bold
        button.style.margin = '2px'; // Adjust margin as needed
        button.style.cursor = 'pointer'; // Set cursor to pointer
        button.style.borderRadius = '5px'; // Set border radius
        button.style.transition = 'background-color 0.3s, box-shadow 0.3s'; // Add transition effect for background color and box-shadow
        button.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Add box shadow for 3D effect

        // Add hover effect
        button.addEventListener('mouseenter', function() {
            button.style.backgroundColor = ' #aed6f1 ';
            button.style.boxShadow = '0px 8px 10px rgba(0, 0, 0, 0.4)'; // Adjust shadow properties for the hover effect
        });

        button.addEventListener('mouseleave', function() {
            button.style.backgroundColor = '#6495ED';
            button.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Reset shadow properties
        });
    }

    moveDropdownStyle(dropdownMenu2);
    moveDropdownStyle(dropdownMenu3);
    moveTextboxStyle(BinInputTextbox);




    /////////////🔄 STYLE CHANGE BUTTONS //////////////////

    function dropdownStyle(button) {
        button.style.backgroundColor = '#d5dbdb'; // Set background color
        button.style.border = '1px solid black'; // Add black border
        button.style.height = '36px';
        button.style.width = '100px';
        button.style.color = 'black'; // Set font color to white
        button.style.padding = '6px 16px'; // Adjust padding as needed
        button.style.textAlign = 'center'; // Center text horizontally
        button.style.textDecoration = 'none'; // Remove underline
        button.style.display = 'inline-block'; // Set display to inline-block
        button.style.fontSize = '14px'; // Set font size
        button.style.fontWeight = 'bold'; // Set font weight to bold
        button.style.margin = '2px'; // Adjust margin as needed
        button.style.cursor = 'pointer'; // Set cursor to pointer
        button.style.borderRadius = '5px'; // Set border radius
        button.style.transition = 'background-color 0.3s, box-shadow 0.3s'; // Add transition effect for background color and box-shadow
        button.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Add box shadow for 3D effect

        // Add hover effect
        button.addEventListener('mouseenter', function() {
            button.style.backgroundColor = '#fbfcfc';
            button.style.boxShadow = '0px 8px 10px rgba(0, 0, 0, 0.4)'; // Adjust shadow properties for the hover effect
        });

        button.addEventListener('mouseleave', function() {
            button.style.backgroundColor = '#d5dbdb';
            button.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Reset shadow properties
        });
    }


    function changeButtonStyle(button) {
        button.style.backgroundColor = '#d5dbdb'; // Set background color
        button.style.border = '1px solid black'; // Add black border
        button.style.height = '24px';
        button.style.width = '100px';
        button.style.color = 'black'; // Set font color to white
        button.style.padding = '6px 16px'; // Adjust padding as needed
        button.style.textAlign = 'center'; // Center text horizontally
        button.style.textDecoration = 'none'; // Remove underline
        button.style.display = 'inline-block'; // Set display to inline-block
        button.style.fontSize = '14px'; // Set font size
        button.style.fontWeight = 'bold'; // Set font weight to bold
        button.style.margin = '2px'; // Adjust margin as needed
        button.style.cursor = 'pointer'; // Set cursor to pointer
        button.style.borderRadius = '5px'; // Set border radius
        button.style.transition = 'background-color 0.3s, box-shadow 0.3s'; // Add transition effect for background color and box-shadow
        button.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Add box shadow for 3D effect

        // Add hover effect
        button.addEventListener('mouseenter', function() {
            button.style.backgroundColor = '#fbfcfc';
            button.style.boxShadow = '0px 8px 10px rgba(0, 0, 0, 0.4)'; // Adjust shadow properties for the hover effect
        });

        button.addEventListener('mouseleave', function() {
            button.style.backgroundColor = '#d5dbdb';
            button.style.boxShadow = '0px 4px 5px rgba(0, 0, 0, 0.2)'; // Reset shadow properties
        });
    }

    dropdownStyle(stateChangeDropdown);
    dropdownStyle(modelChangeDropdown);
    changeButtonStyle(GoneInputTextbox);
    changeButtonStyle(BuildInputTextbox);
    changeButtonStyle(RmaIdInputTextbox);
    changeButtonStyle(shipmentIdinputTextbox);
















    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////






    //////////////////////💾 DOWNLOAD SAN PARTS LIST CSV /////////////////////////////////////////


    function copySpecificColumnsToCSV() {
        const targetColumnsIndices = [2, 3, 4, 7, 8, 9, 10, 11, 12, 13, 14];
        const tableRows = document.querySelectorAll('table.table-bordered.table-striped.table-condensed tbody tr');
        if (!tableRows.length) {
            console.error('Table rows not found');
            return;
        }

        const headers = [
            'Part ID', 'Serial ID', 'Tags', 'Cluster', 'Site', 'Room', 'Bin', 'Model', 'Model MPN',
            'Model APN', 'Model Description', 'Expert Sanitization Verification Ticket',
            'Sanitization Required? (Y/N)', 'Did sanitization occur? (Y/N)', 'Destination', 'Ship By Date'
        ]; // Update with your actual column headers

        const extractedData = [];

        function extractTextContent(element) {
            return element.textContent.trim().replace(/\t|\n/g, '');
        }

        tableRows.forEach(row => {
            let rowDataArray = [];


            ///  CUT OUT TEXT FROM CELLS BEFORE INSERTING INTO EXCEL //////
            targetColumnsIndices.forEach(index => {
                const cellData = row.querySelectorAll('td')[index - 1];
                if (cellData) {
                    let cellText = extractTextContent(cellData);

                    if (index === 2 && cellText.includes('📦')) {
                        cellText = cellText.replace('📦', '');
                    }
                    if (index === 3 && cellText.endsWith('(RMA)')) {
                        cellText = cellText.slice(0, -5); // Remove the last 5 characters
                    }

                    if (index === 4 && cellText.includes('📋SN')) {
                        cellText = cellText.replace('📋SN', '');
                    }
                    if (index === 7 && cellText.includes('📋🗑️')) {
                        cellText = cellText.replace('📋🗑️', '');
                    }
                    if (index === 13 && cellText.includes('Sanitizable: Yes (High Priority)')) {
                        cellText = cellText.replace('Sanitizable: Yes (High Priority)', '');
                    }



                    if (index === 14) {
                        // Assuming the data from index 14 is stored in a variable named 'index14Data'
                        let index14Data = cellText;

                        // Encapsulate the data within double quotes to maintain it as a single cell in CSV
                        index14Data = `"${index14Data.replace(/"/g, '""')}"`;

                        // Push the processed index 14 data into the rowDataArray
                        rowDataArray.push(index14Data);
                    } else {
                        rowDataArray.push(cellText); // Push cell data into the array
                    }
                } else {
                    rowDataArray.push(''); // Push an empty string for an empty cell
                }
            });

            const currentDate = new Date();
            currentDate.setDate(currentDate.getDate() + 2);
            const formattedDate = `${currentDate.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: '2-digit' })}`;

            rowDataArray.push('PASTE TT HERE', 'Y', 'Y', 'VENDOR', formattedDate);
            extractedData.push(rowDataArray);
        });

        const csvContent = [
            headers.join(','),
            ...extractedData.map(row => row.join(','))
        ].join('\n');

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);

        // Append the text from the 2nd column (index 1) of the first row to the filename
        const filename = `${extractedData[0][4]} - Andon Override Exception Parts List - San Required - ${extractedData[0][1]}.csv`;

        link.setAttribute('href', url);
        link.setAttribute('download', filename);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    // Event listener to trigger the CSV download on button click
    document.getElementById('downloadButton').addEventListener('click', copySpecificColumnsToCSV);









    ///////////// 💾 DOWNLOAD CI DATA LIST CSV /////////////////////



    function copySpecificColumnsToCSV3() {
        const tableRows = document.querySelectorAll('table.table-bordered.table-striped.table-condensed tbody tr');
        if (!tableRows.length) {
            console.error('Table rows not found');
            return;
        }

        const extractedData = [];
        const unknownData = [];
        const knownData = [];
        const mpnCountMap = new Map();

        function extractTextContent(element) {
            return element.textContent.trim().replace(/\t|\n/g, '');
        }

        // Create a custom sort function to compare elements based on indice 17 in the Description column
        function customSort(a, b) {
            const descriptionA = a[4]; // Assuming Description is at index 4 in the rowDataArray
            const descriptionB = b[4]; // Assuming Description is at index 4 in the rowDataArray

            if (descriptionA < descriptionB) {
                return -1;
            }
            if (descriptionA > descriptionB) {
                return 1;
            }
            return 0;
        }

        tableRows.forEach(row => {
            const rowDataArray = [];

            rowDataArray.push(null); // 'Asset ID' = null

            // 'Brand' = indice 16
            const brandCellData = row.querySelectorAll('td')[16 - 1];
            let brandCellText = brandCellData ? extractTextContent(brandCellData) : '';

            rowDataArray.push(brandCellText);

            // 'IPN' now uses APN from indice 13
            const apnCellData = row.querySelectorAll('td')[13 - 1];
            const apnCellText = apnCellData ? extractTextContent(apnCellData) : '';
            rowDataArray.push(apnCellText);

            // 'MPN' = indice 11
            const mpnCellData = row.querySelectorAll('td')[11 - 1];
            const mpnCellText = mpnCellData ? extractTextContent(mpnCellData) : '';
            rowDataArray.push(mpnCellText);

            // 'Description' = indice 14
            const descriptionCellData = row.querySelectorAll('td')[14 - 1];
            let descriptionCellText = descriptionCellData ? extractTextContent(descriptionCellData) : '';

            // Extracting content from indice 17 and appending it at the start of Description with ":"
            const additionalInfoCellData = row.querySelectorAll('td')[17 - 1];
            const additionalInfoText = additionalInfoCellData ? extractTextContent(additionalInfoCellData) : '';

            if (additionalInfoText) {
                descriptionCellText = `${additionalInfoText}: ${descriptionCellText}`;
            }

            // Truncate the description after the 5th comma
            if (descriptionCellText) {
                const commaIndex = descriptionCellText.indexOf(',', descriptionCellText.indexOf(',', descriptionCellText.indexOf(',', descriptionCellText.indexOf(',') + 1) + 1) + 1);
                if (commaIndex !== -1) {
                    descriptionCellText = descriptionCellText.slice(0, commaIndex);
                }
            }

            rowDataArray.push(descriptionCellText);

            // 'Serial Number' = indice 3
            const serialNumberCellData = row.querySelectorAll('td')[3 - 1];
            let serialNumberText = serialNumberCellData ? extractTextContent(serialNumberCellData) : '';

            // Remove "(RMA)" from the end of the serial number
            serialNumberText = serialNumberText.replace(/\(RMA\)$/i, '');

            // If the row is not UNKNOWN, exclude serial numbers
            if (!mpnCellText.includes('UNKNOWN')) {
                serialNumberText = null;
            }

            rowDataArray[0] = serialNumberText; // Set 'Asset ID' to Serial Number

            // The rest of the columns with null values
            rowDataArray.push(null, null, null, null, null, null, null, null, null, null);

            // 'Qty' = 1 at the original position
            rowDataArray.splice(10, 0, 1); // Set 'Qty' to 1 at the original position

            // Sort the extractedData array based on indice 17 in the Description column
            extractedData.sort(customSort);




            /////////////////ADDING MPN WHEN UNKNOWN/////////////////////////////////////////////////////////


            // Apply conditions when Brand is "UNKNOWN" or when Brand is empty
            if (brandCellText === 'UNKNOWN' || brandCellText.trim() === '') {
                // Define multiple MPNs for each Brand
                const brandToMPNs = {
                    'Annapurna Labs': [
                        'NT-00138-002', 'K2T-QB', 'K2C-AB', 'K2T-NS1', 'ANP-K2T-QB', 'K2C-A1', 'K2T-QB-1TPM',
                        'ANP-K2C-AB', 'K2T-25X16410010161SS-FHS', 'K2X-N R02', 'MAZN210-000185-001',
                        'K2T-QB-TP1', 'K2X-N', 'K2C-NS-T2', 'K2B-L1', 'K2C-NS1', 'NT-00120',
                        'K2C-25x1600004160X-LPB', 'K2T-25x16410010161TX-FHS', 'K2C-NS', 'K2B-N1',
                        'K2C-ABM-T2', 'PX2-2-T-1 01 R0C V3A', 'PX2-2-T-1', 'PX2-2-T-1A', 'K2V4-N','K2V4-N2', 'K2X-A-32G-T2', 'UNKNOWN_DONGLE'


                        // Add more MPNs here
                    ],

                    'Delta': [
                        'DPS-3000AB-3 A'

                        // Add more MPNs here
                    ],

                    'Accton': [
                        '124-397760-001', '124-700031-001'
                    ],

                    'AMD': [
                        '100-000000091'
                    ],

                    'Eoptolink': [
                        'EOLD-134HG-02-M51', 'EOLQ-131HG-O-02-51'

                    ],

                    'Hynix': [
                        'HMA81GR7CJR8N-WM', 'HMA84GR7CJR4N-XN', 'HMA84GR7DJR4N-XN', 'HMAA8GL7CPR4N-WM', 'HMAA8GR7AJR4N-XN', 'HMAA8GR7CJR4N-XN', 'HMABAGR7A2R4N-XS'

                    ],

                    'INNOLIGHT TECHNOLOGY': [
                        'T-DP4CNT-NWB', 'T-DP4CNT-NWD', 'TR-ZC13T-NWB', 'TR-FC13T-NWD'

                    ],

                    'Intel': [
                        'SPTSBP3CLCSA', 'SPTSBP3CLCZA', 'SPTSLP3SLCWS', 'SPTSHP3PMCWA', 'CD8068904704602'

                    ],

                    'Kingston': [
                        'AM32D429R21D4HJX'
                    ],

                    'Micron': [
                        '36ASF4G72PZ-3G2E7', '36ASF8G72PZ-3G2B2', '9ASF1G72PZ-2G9E1', 'MTA18ASF2G72PDZ-3G2E', 'MTA18ASF2G72PDZ-3G2E1', 'MTA36ASF4G72PZ-3G2E', 'MTA36ASF4G72PZ-3G2E7', 'MTA36ASF4G72PZ-3G2J', 'MTA36ASF4G72PZ-3G2J3', 'MTA36ASF8G72PZ-3G2E1', 'MTA72ASS8G72LZ-2G9J',
                        'MTA36ASF4G72PZ-3G2R1'
                    ],

                    'Samsung': [
                        'M386A8K40BM2-CVF', 'M393A2K43BB3-CVF', 'M393A2K43BB3-CWE', 'M393A4K40BB3-CVF', 'M393A4K40DB3-CWE', 'M393A4K40EB3-CWE', 'M393A8G40AB2-CWE', 'M393A8G40BB4-CWE'
                    ],

                    'Nidec': [
                        'R40W12BGE9-07T841'
                    ],

                    'Avago': [
                        'AFCT-89SFDZ-AZ4'

                    ],

                    'Liteon': [
                        '204-000341'
                    ],

                    // Add more Brand-MPN relationships here
                };

                Object.entries(brandToMPNs).forEach(([brand, mpns]) => {
                    if (mpns.includes(mpnCellText)) {
                        rowDataArray[1] = brand; // Set Brand if the MPN matches any of the MPNs for the brand
                    }
                });
                // Add more conditions for different Brands and their respective MPNs
            }

            ///////////////// MPNS + COUNTRY OF ORIGIN////////////////////////////
            const cooToMPNs = {
                'CHINA': [
                    '070208500-53V-G', '070208500-53V-G', '1LY2BZZ0AM0', 'CS8000-32X-UN-00 (R04)', 'L7048N', 'MTA36ASF4G72PZ-3G2R1',
                    'PS-2471-1L', 'PS-4651-1ZT3', 'QFX3500-48S4Q-AFO', 'QFX3500-RB-ACRB', 'QFX3500-RB-ACRB', 'TR-FC13T-NJC',
                    'TR-QQ13T-NAM', 'TR-QQ13T-NW2', '1064271000', '1064271002', '1064271020', '1064271021', '1064271101',
                    '1064271102', '1064271202', '1064271204', '1064271205', '1064271206', '1064271207', '1064271208', '1064273601',
                    '100-000000091', '100-000000311', '106423-0003', '15216-ATT-LC-12', '15216-ATT-LC-4', '15216-ATT-LC-6',
                    '15216-ATT-LC-8', '15216-ATT-LC-9', '15216-DCU-100', '15216-DCU-550', '15216-DCU-E-200', '15216-MD-40-EVEN',
                    '15216-MD-40-ODD', '15454-OPT-BST', '18ASF1G72PDZ-2G3B1', '18ASF2G72PDZ-2G6D1', '18ASF2G72PDZ-2G6E1',
                    '18ASF2G72PDZ-2G9E', '18ASF2G72PDZ-3G2E', '18ASF2G72PZ-3G2E2', '18JSF1G72PDZ-1G9E2', '18JSF1G72PDZ-1G9P1',
                    '18KSF1G72PDZ-1G6P', '1A21N3500-600-G', '1A42ECC00-600-G', '204-000341', '36ASF4G72PZ-2G6E1', '36ASF4G72PZ-2G9E2', '36ASF4G72PZ-3G2J',
                    '36ASF8G72PZ-3G2B2', '36KSF2G72PZ-1G6N1', '36KSF2G72PZ-1G6P1', '400G DR4-Dummy for Planning', '700-014344-0200',
                    '72ASS8G72LZ-2G6B2', '72ASS8G72LZ-2G6D2', '72ASS8G72LZ-2G9D1', '72ASS8G72LZ-2G9J', '740-011613', '740-013111',
                    '900-2G183-A800-001', '900-2G183-A800-001', '900-2G503-0010-000', '965-2G506-A831-000', '9ASF1G72PZ-2G6D1',
                    '9ASF1G72PZ-2G9E1', 'AFBR-79E4Z-D-JU1', 'AFCT-89SFDZ-AZ1', 'AFCT-89SFDZ-AZ2', 'AFCT-91DRDDZ-AZ1',
                    'AM12D426R19D8HABX1', 'AM16D3L16R11D4HB', 'AM16D3L16R11D4HB1', 'AM16D424R17D8MAM', 'AM16D426R19D8HAX1',
                    'AM32D424R17D4MAM', 'AM32D426R19D4HCX', 'AM32D426R19D4HJX', 'AM32D426R19D4MEX', 'AM32D426R19D4NAPX1',
                    'AM32D429R21D4MEX', 'AM4D424R17S8HAX1', 'AM4D424R17S8MBM', 'AM4D426R19S8HBX1', 'AM8D424R17D8HBX1',
                    'AM8D424R17D8MBM', 'AM8D424R17D8MBM', 'AM8D426R19D8HBX1', 'AQOL-BCQ4-EDMA-0776', 'AQOL-BCQ4-EDMA-0776',
                    'AQOL-LCQ4-EDMA-0662', 'AQOL-LCQ4-EDMA-0662', 'AQOL-LCQ4-EDMA-0791', 'AQOMBBQ4EDMA0694', 'AQOMBBQ4EDMA0918',
                    'AQOMLBQ4EDMA0730', 'AQOM-LBQ4-EDMA-0730', 'AQOMNN85ADLN0534', 'AQOM-NN85-ADLN-0534', 'AQOMNN85ADLN0688',
                    'AQP1AN03EDMA0819', 'AQP1AN1AEDMA0819', 'AQP1AN2AEDMA0819', 'AQPLBCQ4EDMA0891', 'AQPLLCQ4EDMA0894',
                    'CFP-100GBASE-ER4', 'CFP-100GBASE-LR4', 'CFP-GEN2-100GBASE-LR4', 'CFP-GEN2-CGE-ER4', 'CV80E12BS2ND5-07T02',
                    'DF40-GE-LR4-111-00F', 'DFPD0856B2UY018', 'DFPD0856B2UY033', 'DPS-1200AB-14 A S2F', 'DPS-1200AB-14 B',
                    'DPS-1200AB-14 B', 'DPS-1200AB-14 D', 'DPS-1200AB-15 B', 'DPS-1200AB-15 B', 'DPS-1200AB-15 C', 'DPS-1600CB-F',
                    'DPS-3000AB-3 A', 'DPS-3000AB-3 A', 'DPS-3000AB-3 A_00F', 'DPS-3000AB-3 A_01F', 'DPS-500AB-16 C', 'DPS-650XB-2 J',
                    'DPS-650XB-2G', 'DPS-750AB-26 A', 'DQ5DC88P043', 'DS-SFP-FC8G-SW', 'EOLD-134HG-02-M5', 'EOLD-134HG-02M51',
                    'EOLD-134HG-02-M51', 'EOLD-134HG-02-M55', 'EOLD-134HG-02-M56', 'EOLD-134HG-02-M57', 'EOLD-134HG-02-M59',
                    'EOLD-134HG-02-M5A', 'EOLD-134HG-02-N51', 'EOLD-164HG-10-5', 'EOLQ-131HG-O-02-5', 'EOLQ-131HG-O-02-51', 'EOLQ-161HG-10-L', 'EX-SFP-1GE-LX', 'EX-SFP-1GE-SX', 'EX-SFP-1GE-T', 'EX-SFP-1GE-T', 'EX-UM-2X4SFP', 'FCLF-8521-3',
                    'FCMJ-8521-3', 'FSE023-4Z0G', 'FSE023-931G', 'FSE023-931G', 'FSE037-4Z1G', 'FSH005-930G', 'FSH005-930G',
                    'FSH005-930G', 'FSH005-B60G', 'FTL410QD3C', 'FTL410QD3C', 'FTL410QE2C', 'FTL410QE2C', 'FTLF1318P2BTL',
                    'FTLF1318P2BTL', 'GFC0412SS-03', 'GFC0412SS-03', 'HL13B5CP10-L2', 'HMA41GR7AFR8N-TF', 'HMA41GR7AFR8N-UH',
                    'HMA42GR7AFR4N-TF', 'HMA451R7AFR8N-UH', 'HMA815R7AFR8N-VK', 'HMA81GR7AFR8N-VK', 'HMA81GR7CJR8N-VK',
                    'HMA81GR7CJR8N-WM', 'HMA82GR7AFR8N-UH', 'HMA82GR7AFR8N-VK', 'HMA82GR7CJR8N-VK', 'HMA82GR7CJR8N-WM',
                    'HMA82GR7JJR8N-WM', 'HMA84GR7AFR4N-UH', 'HMA84GR7AFR4N-VK', 'HMA84GR7CJR4N-VK', 'HMA84GR7CJR4N-WM',
                    'HMA84GR7CJR4N-XN', 'HMA84GR7DJR4N-XN', 'HMA84GR7JJR4N-VK', 'HMA84GR7JJR4N-WM', 'HMA84GR7MFR4N-TF',
                    'HMA84GR7MFR4N-UH', 'HMAA4GR7AJR4N-XN', 'HMAA8GL7AMR4N-VK', 'HMAA8GL7CPR4N-WM', 'HMAA8GL7MMR4N-UH',
                    'HMAA8GR7A2R4N-VN', 'HMAA8GR7AJR4N-XN', 'HMABAGR7A2R4N-XS', 'HMABAGR7A4R4N-VN', 'HMCG78MEBRA',
                    'HMT151R7AFP4C-H9', 'HMT31GR7CFR4A-H9T', 'HMT351R7CFR8A-H9T', 'HMT351R7CFR8A-H9T', 'HMT351R7CFR8C-PBT',
                    'HMT351U7CFR8A-H9', 'HMT41GR7AFR8A-PB', 'HMT41GR7AFR8C-PB', 'HMT41GR7AFR8C-RD', 'HMT41GR7BFR8A',
                    'HMT41GR7BFR8C-PB', 'HMT41GR7BFR8C-RD', 'HMT41GR7DFR8C-RD', 'HMT41GU7AFR8A-PB', 'HMT41GU7BFR8A-PB',
                    'HMT42GR7AFR4A-PB', 'HMT42GR7AFR4C-RD', 'HMT42GR7BFR4A-PB', 'HMT42GR7BFR4C-RD', 'HMT42GR7DFR4A-PB',
                    'HMT42GR7MFR4A-H9T', 'HMT451R7AFR8C-PB', 'HMT84GL7BMR4A-PB', 'JNP-QSFP-100G-CWDM', 'JNP-QSFP-100G-LR4',
                    'JNP-QSFP-100G-LR4', 'JNP-QSFP-100G-PSM4', 'JNP-QSFP-100G-SR4', 'JNP-QSFP-40GE-IR4', 'JNP-QSFP-40G-LR4',
                    'JNP-QSFP-40G-LX4', 'LQ210CR-CPA2', 'LQ210CR-CPA7', 'LUX42604B0-FIB3-DR2', 'LUX42604BO', 'LUX42604BP-LC0180',
                    'M386A8K40BM2-CVF', 'M386A8K40BMB-CRC', 'M386A8K40CM2-CTD', 'M393A1543BB1-CTD', 'M393A1G43DB0-CPB',
                    'M393A1G43EB1-CRC', 'M393A1K43BB1-CTD', 'M393A2G40DB0-CPB', 'M393A2K43BB1-CRC', 'M393A2K43BB1-CTD', 'M393A2K43BB3-CVF', 'M393A2K43CB2-CTD',
                    'M393A4K40BB0-CPB', 'M393A4K40BB1-CRC', 'M393A4K40BB2-CTD', 'M393A4K40BB3-CVF', 'M393A4K40CB1-CRC',
                    'M393A4K40CB2-CTD', 'M393A5143DB0-CRC', 'M393A5143EB0-CRC', 'M393A8G40BB4-CWE', 'M393A8K40B22-CWD',
                    'M393AAG40M32-CYF', 'M393AAG40M3B-CYF', 'M393AAK40B42-CWD', 'M393ABG40M52-CYF', 'M393ABG40M5B-CYF',
                    'M393B1G73BH0-CK0', 'M393B1G73BH0-YK0', 'M393B1G73DB0-CMA', 'M393B1G73DB0-YK0', 'M393B1G73EB0-CMA',
                    'M393B1G73EB0-YK0', 'M393B1G73QH0-CMA', 'M393B1G73QH0-YK0', 'M393B1K70CH0-YH9', 'M393B1K70DH0-CK0',
                    'M393B1K70DH0-YH9', 'M393B2G70BH0-CH908', 'M393B2G70BH0-CK0', 'M393B2G70BH0-YH9', 'M393B2G70BH0-YK0',
                    'M393B2G70DB0-CMA', 'M393B2G70DB0-YK0', 'M393B2G70EB0-CMA', 'M393B2G70EB0-YK0', 'M393B2G70QH0-YK0',
                    'M393B5273CH0-CH9', 'M393B5273DH0-YH908', 'M393B5773DH0-YH9', 'MC1200B4-3-4R-02', 'MC1200B4-3-4R-02 REV D2',
                    'MC1200B4-3-4R1-02', 'MTA36ASF4G72PZ-3G2R1', 'MTA36ASF8G72PZ-3G2F1', 'MTC10F1084S1RC48BA', 'PS-2122-7Q REV X3',
                    'PS-2122-8L', 'PS-2122-8L', 'PS-2122-8L1', 'PS-2162-1Q1', 'PS-2162-1Q1', 'PS-2162-1Q1', 'PS-2162-1Q1',
                    'PS-2322-1ADU-LF', 'PS-2322-1ADU-LF', 'PS-2322-1ADU-LF', 'PS-2322-1ADU-LF', 'PS-2322-1ADU-LF', 'PS-4651-1HY1-LF',
                    'PS-4651-1ZT3', 'QFX-QSFP-40G-ESR4', 'QFX-QSFP-40G-SR4', 'QSFP-100GB4FDRA', 'QSFP-100GB4FDRB', 'R1165-F8001-01',
                    'R1165-F8002-01', 'R40W12BS1NL9-07T02', 'S800E005L', 'S800E005L', 'SFP-1GE-LX', 'SPP-10E-LR-CDFP', 'SPQ-10E-LR-CDFL',
                    'SPQ-CE-CL-CDFM', 'SPQ-CE-LR-CDFA', 'SPTSBP3CLCAS', 'T-DP4CNT-NW5', 'T-DP4CNT-NW5', 'T-DP4CNT-NWB', 'T-DP4CNT-NWD',
                    'TP-VC13T01P8-N00', 'TP-VC13T02P7-N00', 'TR-FC13T-NAM', '07020JR00-53V-G', '1LB9BZZ000N', '700-014344-0000', '740-073146', 'CS8000-32X-UN-00', 'CS8000-32X-UN-00 301-000113-001 R1165-F0001-01',
                    'DPS-1200AB-20 C', 'DPS-4200AB A S0F', 'DPS-4200AB A S1F', 'JW657A', 'LB9', 'PS-2422-1ADU', 'PS-2422-1ADU Rev.X2', 'PS-2422-1ADU Rev.X7', 'PS-4801-2FR',
                    'XRV-PSU1-770W', '39747387', '1002971151', '1002971251', '1064281107', '1064281111', '1837042024', '1837042123', '1837048257', '2015912003', '2015912004', '2015912005',
                    '2179891005', '2179891211', '2179891213', '02011AY02-600-G', '02011AY03-600-G', '02011B102-600-G', '02011EU00-600-G', '02011MR00-600-G', '100-000000038', '100-000000349',
                    '100-000000905-00', '100297-1201', '100-341B', '100-499B', '100-505993', '100-506015', '110099-010MB', '110099-014MB', '110100-010MB', '110100-012MB', '110100-028MB',
                    '110214-005M', '110319-008MB', '110319-010MB', '110319-012MB', '110319-014MB', '110319-016MB', '110319-018MB', '11MBZZZ020R', '11MBZZZ020U', '1395A2491615', '1395A2491618',
                    '1395A2618401', '1395A2618704', '1395A2689901', '1395A2725601', '144ASQ16G72PSZ-2S6E1', '144ASQ16G72PSZ-2S6E1S', '144ASQ16G72PSZ-2S6G1', '144ASQ32G72PSZ-3S2B', '15216-ATT-LC-10',
                    '15216-ATT-LC-15', '15216-ATT-LC-3', '15216-ATT-LC-5', '15216-ATT-LC-7', '15216-EF-40-EVEN', '15216-EF-40-ODD', '15216-MD-48-CM', '15216-MD-ID-50', '15454-40ME-MXP-C',
                    '15454-M-100G-ME-C', '15454-M-10X10G-LC', '15454-M6-AC2', '15454-M6-DC', '15454-M6-ECU', '15454-OPT-AMP-C', '15454-OPT-EDFA-17', '15454-OPT-PRE', '16ATF2G64HZ-2G6E3', '177628-001',
                    '18ASF1G572PDZ-2G3111', '18ASF1G572PDZ-2G6221', '18ASF1G572PDZ-2G6331', '18ASF1G72PDZ-2G1A1', '18ASF1G72PDZ-2G3A2', '18ASF1G72PDZ-2G6E1', '18ASF1G72PZ-2G4', '18ASF2G72PDZ-2G3A1', '18ASF2G72PDZ-2G3B1',
                    '18ASF2G72PDZ-2G6B1', '18ASF2G72PDZ-2G9J', '18ASF2G72PDZ-3G2J', '18ASF2G72PZ-3G2J', '18ASF4G72PDZ-2G9B', '18ASF4G72PDZ-2G9E',
                    '18ASF4G72PDZ-3G2B', '18ASF4G72PDZ-3G2E', '18ASF4G72PZ-2G9B1', '18ASF4G72PZ-2G9E', '18ASF4G72PZ-3G2B', '18ASF4G72PZ-3G2E',
                    '18JSF1G72PDZ', '18JSF1G72PDZ-1G6', '18JSF1G72PDZ-1G6N1', '18JSF1G72PZ-1G6P1', '18JSF1G72PZ-1G9E1', '18JSF1G72PZ-1G9P1',
                    '18JSF51272AZ-1G4', '18KSF1G72AZ-1G6E1', '18KSF1G72AZ-1G6P1', '18KSF1G72HZ-1G6P1', '18KSF1G72PDZ', '18KSF1G72PDZ-1G6E1',
                    '18KSF51272AZ-1G6', '18KSF51272PZ-1G4D1', '18KSF51272PZ-1G4M1', '1A32FWY00-600-G', '1A32MWQ00-600-G', '1A32Q9Q00-600-G',
                    '1A32Q9S00-600-G', '1A32U8900-600-G', '1A32WWU00-600-G', '1A32WWU00-600-G', '1A427JY00-600-G', '1A5227100-600-G', '249675-001',
                    '249676-001', '260741-001', '2AS2VMA0030', '2LS4LMA0000', '300701-001', '31S1DMB0020', '31S1GMB0020', '31S1PMB0000', '31S1VMB0040',
                    '31S1VMB0050', '31S1VMB0080', '31S2MMB00B0', '31S2MMB00D0', '321851-001', '34-00-00204-R', '34-00-00205-R', '34-00-00251-R',
                    '34-00-00279-R', '34-00-00287-R', '34-02-00020-R', '34-02-00021-R', '34-02-00025-R', '34-02-00026-R', '34-02-00040-R',
                    '34-02-00041-R', '34-03-00019-R', '34-03-00043-R', '34-03-00045-R', '34-03-00046-R', '34-03-00049-R', '34-03-00070-R',
                    '34-03-00077-R', '34-03-00079-R', '34-04-00020-R', '361960-001', '36ASF2G72PZ-2G1A2', '36ASF2G72PZ-2G4', '36ASF4G72PZ-2G1',
                    '36ASF4G72PZ-2G3A1', '36ASF4G72PZ-2G3B1', '36ASF4G72PZ-2G3B1C', '36ASF4G72PZ-2G6D1', '36ASF4G72PZ-2G6E1C', '36ASF4G72PZ-2G6H1',
                    '36ASF4G72PZ-2G9J', '36ASF4G72PZ-3G2E7', '36ASF8G72PZ-2G9B', '36ASF8G72PZ-2G9E', '36ASF8G72PZ-3G2E', '36JSF1G72PZ-1G6M1',
                    '36JSF2G72PZ-1G9E1', '36JSF2G72PZ-1G9N1', '36JSF2G72PZ-1G9P1', '36JSZF51272PY-1G4', '36KSF1G72PZ-1G4', '36KSF1G72PZ-1G4M1',
                    '36KSF2G72PZ-1G6', '370780-001', '378915-001', '398707-051', '398708-061', '39C947382E', '405477-061', '40A1G8SA-062E:E',
                    '416107-001', '416471-001', '416474-001', '432668-001', '451U7AFR8APB', '455263-061', '466436-061', '500202-061', '501538-001',
                    '5411C4020030', '5411C4020031', '5411C4020032', '5411C4020033', '5411C4020034', '5541843-B', '606426-001', '647897-B21',
                    '675-23587-0000-QS1', '684034-001', '692-2G504-0200-010', '699-2G183-0205-RC1', '699-2G503-0201-QS1', '700-014344-0000',
                    '700-014344-0000', '708637-B21', '726719-B21', '726719-S21', '72ASS16G72LZ-3G2B', '72ASS16G72LZ-3G2B3', '72ASS16G72PSZ-3S2B',
                    '72ASS8G72LZ-2G3A1', '72ASS8G72LZ-2G3B2', '72ASS8G72LZ-2G6D2Q', '72ASS8G72LZ-2G6D2S', '72ASS8G72LZ-2G9J2', '72ASS8G72PSZ-2S6E',
                    '72ASS8G72PSZ-2S6E1', '72ASS8G72PSZ-2S6G1', '72KSZS4G72LZ-1G6E2', '72KSZS4G72PZ-1G4E', '740-031851', '753221-B21',
                    '900-12055-0010-000', '900-12055-0020-000', '900-22080-0000-000', '900-2G133-A840-000', '900-2G133-A840-100', '900-2G183-0000-001',
                    '900-2G183-0000-001', '900-2G183-0000-001', '900-2G183-A800-002', '900-2G402-0000-000', '900-2G414-0000-000', '900-2G500-0000-000',
                    '900-2G503-0000-000', '900-2G503-0000-000', '900-2G600-0000-001', '900-2H403-0000-000', '965-2G506-0031-200', '965-2G510-A831-000',
                    '9965516-420.A00LF', '9ASF2G72PZ-2G9E', '9ASF2G72PZ-3G2B', '9ASF2G72PZ-3G2E', '9ASF51272PZ-2G1A2', '9ASF51272PZ-2G3A2',
                    '9ASF51272PZ-2G3B1', '9ASF51272PZ-2G6E1', '9J5WF', '9J5WF', '9J5WF', '9J5WF', '9JSF51272PZ-1G6E1', '9KSF25672PZ-1G6K1',
                    '9KSF51272AZ-1G6E1', '9KSF51272PZ-1G6', '9KSF51272PZ-1G6P1', 'AM12D426R19D5HABPX1', 'AM32D426R19D4NAX1', 'AM32D429R21D4HJX',
                    'AM372D3LD8P13C', 'AM372D3LD8P13C9H', 'AM472D3LD4P13C9HC', 'AM4D3L16E11S8HA', 'AM4D3L16R11S8HA', 'AM572D3LD4P13C9HM', 'AM8D3L16E11D8HA',
                    'AM8D3L16R11D8HBX', 'AM8D3L16SE11D8NE1', 'AM8D424R17D8HAX', 'AM8D429R21S8HCX1', 'ATR38AGML33', 'ATR3GAGMW19', 'B55.00501.D010',
                    'C13-C14-3M', 'C14-C15-2M', 'CBL-EX-PWR-C13-US', 'CCAEIA-DS002-A080-C1B', 'CCAEIA-DS002-A160-C5B',
                    'CPAK-100G-LR4', 'CS1200-MN3-440', 'CS1200-MN3-441', 'CT25672AA667M18FH', 'CT25672AA667M18FH', 'CT51272BA1339',
                    'CT51272BA1339', 'CX1A0002R0008.003M', 'CX1A0008R6D04-2.1M', 'DA40PSA10BS', 'DA40PSA15BS', 'DA40PSA20BS',
                    'DA40PSA25BS', 'DA40PSA30BS', 'DBPG0428B2UP005B-1', 'DFPB0956B2HY007', 'DFPD0856B2UY005', 'DFPD0856B2UY022',
                    'DFPD0856B2UY063', 'DFPK0456B2GY0C6', 'DFPQ0456B2GY020', 'DPS-1200AB-20 C', 'DPS-3000AB-3 A_02F', 'DPS-3000AB-3 A_03F',
                    'DPS-4200AB A   S5', 'DPS-4200AB A  S3', 'DPS-4200AB A 00', 'DPS-4200AB A S2F', 'DPS-4200AB A S4', 'DQ5DC88P045',
                    'DX42BSA15YA', 'DX42BSA20YA', 'DX42BSA25YA-1', 'EX4500-PWR1-AC-FB', 'EX745', 'EX8200-PWR-AC2K', 'FA122A08-H01',
                    'FCBG110SD1C02B', 'FCBG110SD1CB5B', 'FIBRETESTER', 'FPR-C9300-AC', 'FTLF1318P3BTL', 'FTLX1475D3BCL', 'FTLX1672D3BCL',
                    'G484D', 'GFM0412SS-08', 'GFM0412SS-08FX8', 'GFM0412SS-10J1T', 'GG8067402570204', 'GG8067402570603', 'HMA41GR7AFR4N-UH',
                    'HMA41GR7BJR8N-VK', 'HMA41GR7MFR8N-TF', 'HMA42GR7AFR4N-UH', 'HMA42GR7MFR4N-TF', 'HMA451R7AFR8N-TF', 'HMA815R7MFR8N-UH',
                    'HMA82GR7CJR4N-WM', 'HMA82GR7CJR4N-XN', 'HMA82GR7DJR4N-XN', 'HMA82GR7DJR8N-WM', 'HMA82GR7JJR8N-VK', 'HMA82GR7MFR8N-ASR',
                    'HMA82GS6CJR8N-VK', 'HMA84GL7AFR4N-VK', 'HMA84GR7AFR4N-UHC2', 'HMA84GR7BJR4N-UHC', 'HMA84GR7DJR4N-WM', 'HMAA4GR7AJR8N-WM',
                    'HMAA4GR7AJR8N-XN', 'HMAA4GR7CJR4N-XN', 'HMAA4GR7MJR8N-WM', 'HMAA4GR7MJR8N-XN', 'HMAA8GL7CPR4N-XN', 'HMAA8GR7CJR4N-WM',
                    'HMAA8GR7MJR4N-WM', 'HMAA8GR7MJR4N-XN', 'HMABAGL7A4R4N-VN', 'HMABAGL7ABR4N-WM', 'HMABAGL7ABR4N-XN', 'HMABAGL7CBR4N-XN',
                    'HMABAGR7C4R4N-WR', 'HMABAGR7C4R4N-XS', 'HMAG74EXNRA', 'HMAG84DXNRB', 'HMAG84EXNRA', 'HMAG94DXNRB', 'HMAT14JXSRB',
                    'HMCG78AEBRA', 'HMCG78AEBRA107N', 'HMCG78AEBRA110N', 'HMCG78AEBRA115N', 'HMCG78AEBRA168N', 'HMCG78AGBRA',
                    'HMCG78AGBRA186N', 'HMCG78AGBRA188N', 'HMCG78AGBRA190N', 'HMCG78AGBRA191N', 'HMCG78BHBRA284N', 'HMCG78BHBRA288N',
                    'HMCG78BHBRA290N', 'HMCG78BHBRA292N', 'HMCG78MEBRA107N', 'HMCG78MEBRA113N', 'HMCG78MEBRA115N', 'HMCG78MEBRA174N',
                    'HMCG84AEBQA', 'HMCG84AEBRA', 'HMCG84AEBRA107N', 'HMCG84AEBRA110N', 'HMCG84AEBRA115N', 'HMCG84AEBRA168N',
                    'HMCG84AGBQA', 'HMCG84AGBRA', 'HMCG84AGBRA186N', 'HMCG84AGBRA188N', 'HMCG84AGBRA190N', 'HMCG84AGBRA191N',
                    'HMCG84BHBRA284N', 'HMCG84BHBRA288N', 'HMCG84BHBRA290N', 'HMCG84BHBRA292N', 'HMCG84MEBRA', 'HMCG84MEBRA107N',
                    'HMCG84MEBRA113N', 'HMCG84MEBRA115N', 'HMCG84MEBRA174N', 'HMCG88AEBRA', 'HMCG88AEBRA107N', 'HMCG88AEBRA110N',
                    'HMCG88AEBRA115N', 'HMCG88AEBRA168N', 'HMCG88AGBRA', 'HMCG88AGBRA186N', 'HMCG88AGBRA188N', 'HMCG88AGBRA190N',
                    'HMCG88AGBRA191N', 'HMCG88BHBRA284N', 'HMCG88BHBRA288N', 'HMCG88BHBRA290N', 'HMCG88BHBRA292N', 'HMCG88MEBRA',
                    'HMCG88MEBRA107N', 'HMCG88MEBRA113N', 'HMCG88MEBRA115N', 'HMCG88MEBRA174N', 'HMCG94AEBQA', 'HMCG94AEBRA',
                    'HMCG94AEBRA102N', 'HMCG94AEBRA108N', 'HMCG94AEBRA109N', 'HMCG94AEBRA123N', 'HMCG94AGBQA', 'HMCG94AGBRA',
                    'HMCG94AGBRA177N', 'HMCG94AGBRA179N', 'HMCG94AGBRA181N', 'HMCG94AGBRA182N', 'HMCG94BHBRA275N', 'HMCG94BHBRA279N',
                    'HMCG94BHBRA281N', 'HMCG94BHBRA283N', 'HMCG94MEBQA', 'HMCG94MEBRA', 'HMCG94MEBRA109N', 'HMCG94MEBRA112N',
                    'HMCG94MEBRA121N', 'HMCG94MEBRA123N', 'HMCGJ8MEBRB', 'HMCGJ8MEBRB215N', 'HMCGJ8MEBRB216N', 'HMCGJ8MEBRB218N',
                    'HMCGJ8MEBRB220N', 'HMCGJ8MGBRB', 'HMCGJ8MGBRB223N', 'HMCGJ8MGBRB227N', 'HMCGJ8MGBRB228N', 'HMCGJ8MGBRB231N',
                    'HMCGM4MEBRB175N', 'HMCGM4MEBRB233N', 'HMCGM4MEBRB235N', 'HMCGM4MEBRB237N', 'HMCGY4MGBRB223N', 'HMCGY4MGBRB227N',
                    'HMCGY4MGBRB228N', 'HMCGY4MGBRB231N', 'HMCGY8MGBRB223N', 'HMCGY8MGBRB227N', 'HMCGY8MGBRB228N', 'HMCGY8MGBRB231N',
                    'HMCT04AEERA', 'HMCT04AEERA128N', 'HMCT04AEERA131N', 'HMCT04AEERA135N', 'HMCT04AEERA138N', 'HMCT04AGERA',
                    'HMCT04AGERA195N', 'HMCT04AGERA197N', 'HMCT04AGERA199N', 'HMCT04AGERA200N',
                    'HMCT04BHERA293N', 'HMCT04BHERA297N', 'HMCT04BHERA299N', 'HMCT04BHERA301N', 'HMCT04MEERA', 'HMCT14AEERA', 'HMCT14AEERA144N', 'HMCT14AEERA147N', 'HMCT14AEERA152N', 'HMCT14AEERA155N',
                    'HMCT14AGERA', 'HMCT14AGERA204N', 'HMCT14AGERA206N', 'HMCT14AGERA208N', 'HMCT14AGERA209N', 'HMCT14BHERA302N', 'HMCT14BHERA306N', 'HMCT14BHERA308N', 'HMCT14BHERA310N', 'HMCT14MEERA',
                    'HMT151R7AFP4', 'HMT151R7BFR4C-H9', 'HMT151R7TFR4A-H9', 'HMT151R7TFR4C-H9', 'HMT31GR7BFR4C-PBD', 'HMT31GR7BFR4C-PBT', 'HMT31GR7CFR4C-PB', 'HMT31GR7EFR4A-PBT', 'HMT31GR7EFR4C-RDT',
                    'HMT325R7CFR8A-H9', 'HMT325R7CFR8A-H9T', 'HMT351R7BFR4A-H9T', 'HMT351R7BFR8C-H9', 'HMT351R7CFR8C-H9T', 'HMT351R7EFR8A-PBT', 'HMT41GA7DFR8A-PB', 'HMT41GR7AFR4C-RD', 'HMT41GR7DFR8A-P',
                    'HMT42DGR7AFR4A', 'HMT42GR7AFR4A-H9', 'HMT42GR7AFR4C-PB', 'HMT42GR7MFR4C-PB', 'HMT42GR7MFR4C-PBT', 'HMT451R7AFR8A-PB', 'HMT451R7BFR8A-PB', 'HMT451U7AFR8A-PB', 'HMT451U7BFR8A-PB',
                    'HMT84GL7AMR4A-PB', 'HMT84GR7MMR4A-H9', 'HPO-398645-001', 'HPO-416472-001', 'HPO-416473-001', 'HPO-467654-001', 'HPO-501534-001', 'HPO-501536-001', 'HYMP151P72CP4-S5', 'HYMP151P72CP4-S6',
                    'JNP-100G-AOC-5M', 'JNP-FAN-1RU-BB', 'JNP-PWR1600-AC-BB', 'JPSU-1100-AC- AFO', 'JPSU-1600-C-AC-AFO', 'JPSU-400W-AC', 'JW657A', 'K374T', 'K4AAG085WA-BCWE', 'KAZ-HMT41GR7AFR8A-PB',
                    'KVR16LE11/8KF', 'KVR16LR11D8/8KF', 'KVR16LR11S4/8KF', 'KVR16R11S4/8KF', 'KVR667D2D4P54G', 'KVR667D2D8P52G', 'KVR667D2E5 2G', 'KVR667D2N5 2G', 'LCLC-SM-30M', 'M386A8K40CM2-CVF',
                    'M386A8K40CM2-CVF', 'M386A8K40DM2-CWE', 'M386AAG40AM3-CWE', 'M386AAG40BM3-CWE', 'M386B4G70DM0-YK0', 'M386B4G70DM0-YK03', 'M391B5273DH0-YH9', 'M393A1543BB1-CRC', 'M393A1G40DB0-RC',
                    'M393A1G43DB1-CRC', 'M393A1K43DB1-CWE', 'M393A1K43FB2-CWE', 'M393A2G40DB0-CRC', 'M393A2K40DB3-CWE', 'M393A2K40EB3-CWE', 'M393A2K43CB2-CVF', 'M393A2K43CB2-CVF', 'M393A2K43DB2-CVF',
                    'M393A2K43DB3-CWE', 'M393A2K43FB3-CWE', 'M393A4G40AB3-CWE', 'M393A4G40BB3-CWE', 'M393A4G40CB3-CWE', 'M393A4G43AB3-CVF', 'M393A4G43AB3-CWE', 'M393A4K40BB0-CRC', 'M393A4K40CB2-CVF',
                    'M393A4K40CB2-CVF', 'M393A4K40DB2-CVF', 'M393A4K40EB3-CWE', 'M393A5143EB1-CRC', 'M393A8G40AB2-CVF', 'M393A8G40CB4-CWE', 'M393A8G40MB2-CVF', 'M393B1G70DB0-CMA', 'M393B1G70QH0-CMA',
                    'M393B1G73BH0-YH9', 'M393B1G73QH0-CK0', 'M393B1GB73BHO-YKO', 'M393B1K70', 'M393B1K70CH0-CH9', 'M393B1K70DH0-CH9', 'M393B1K70DH0-CK9', 'M393B1K70DH0-CMA', 'M393B1K70DH0-YH908',
                    'M393B1K70DH0-YK0', 'M393B2970BH4-YK008', 'M393B2G70BH0-YK908', 'M393B2G70QH0-CK0', 'M393B4G70BM0-YH9', 'M393B4G70DM0-YK0', 'M393B5170EH1-CH9', 'M393B5170FH0', 'M393B5170FH0-YH9',
                    'M393B5170FHO-CH9', 'M393B5270DH0-YH9', 'M393B5273CH0-YH9', 'M393B5273DH0-CH9', 'M393B5273DH0-CH909', 'M393B5273DH0-CK0', 'M393B5273DH0-CK008', 'M393B5273DH0-YH9', 'MAZN80010BM00-216-G',
                    'MAZN800202900-738-G', 'MAZN800303J00-220-G', 'MAZN800303N00-026-G', 'MAZN80030450-220-G', 'MAZN800304900-026-G', 'MAZN800304L00-026-G', 'MAZN800305100-026-G', 'MAZN800305800-220-G',
                    'MB-A9SD6F-10STP', 'MB-A9SHL6F(-N)', 'MEM-DR340L-CL02-ER13', 'MEM-DR340L-SL06-ER13', 'MEM-DR380L-HL04-ER13', 'MEM-DR380L-HL05-ER13', 'MEM-DR380L-SL10-ER16', 'MEM-DR380L-SL12-ER16',
                    'MG5YT', 'MG5YT', 'MPC4E-3D-2CGE-8XGE-RB', 'MS-MPC-128G', 'MT18HTF25672AY-667G1', 'MT18KSF1G72PDZ1G', 'MT18KSF51272PDZ-1G6K', 'MT36KSF2G72PZ-1G6E1L', 'MT40A2G8JC-062E:E',
                    'MTA18ASF2G72PZ-3G2R1', 'MTA18ASF4G72PDZ-3G2F1', 'MTA18ASF4G72PZ-3G2F1', 'MTA36ASF8G72LZ-3G2B1', 'MTA72ASS16G72LZ-3G2F1', 'MTA72ASS8G72LZ-3G2R2', 'MTA9ASF2G72PZ-3G2F1', 'MTC10F1084S1RC48BA1001',
                    'MTC10F1084S1RC48BA1003', 'MTC10F1084S1RC48BA1004', 'MTC10F1084S1RC48BA100B', 'MTC10F1084S1RC48BA19', 'MTC10F1084S1RC48BA19JH', 'MTC10F1084S1RC48BA19NG', 'MTC10F1084S1RC48BA19NI',
                    'MTC10F1084S1RC48BA19PI', 'MTC10F1084S1RC56BG1', 'MTC10F108YS1RC48BB1', 'MTC10F108YS1RC48BB1001', 'MTC10F108YS1RC48BB1002', 'MTC10F108YS1RC48BBZ', 'MTC10F108YS1RC56BB1', 'MTC20F1045S1RC48BA',
                    'MTC20F1045S1RC48BA2003', 'MTC20F1045S1RC48BA2004', 'MTC20F1045S1RC48BA2006', 'MTC20F1045S1RC48BA200C', 'MTC20F1045S1RC48BA2JG', 'MTC20F1045S1RC48BA2JH', 'MTC20F1045S1RC48BA2NG',
                    'MTC20F1045S1RC48BA2PI', 'MTC20F2085S1RC48BA', 'MTC20F2085S1RC48BA1002', 'MTC20F2085S1RC48BA1003', 'MTC20F2085S1RC48BA1004',
                    'MTC20F2085S1RC48BA100B', 'MTC20F2085S1RC48BA19', 'MTC20F2085S1RC48BA19JH', 'MTC20F2085S1RC48BA19NG', 'MTC20F2085S1RC48BA19NH', 'MTC20F2085S1RC48BA19PI', 'MTC20F2085S1RC52BG1', 'MTC20F2085S1RC56BD1001', 'MTC20F2085S1RC56BD1001',
                    'MTC20F2085S1RC56BD1002', 'MTC20F2085S1RC56BD1002', 'MTC20F2085S1RC56BD100A', 'MTC20F2085S1RC56BD100A', 'MTC20F2085S1RC56BD100B', 'MTC20F2085S1RC56BD100B', 'MTC20F2085S1RC56BG1', 'MTC20F2085S1RC64BD1001', 'MTC20F2085S1RC64BD1003', 'MTC20F2085S1RC64BD100A',
                    'MTC20F2085S1RC64BD100B', 'MTC40F2046S1RC48BA', 'MTC40F2046S1RC48BA1002', 'MTC40F2046S1RC48BA100B', 'MTC40F2046S1RC48BA100C', 'MTC40F2046S1RC48BA1FI', 'MTC40F2046S1RC48BA1IG', 'MTC40F2046S1RC48BA1IH', 'MTC40F2046S1RC48BA1MH', 'MTC40F2046S1RC52BG1', 'MTC40F2046S1RC56BD1',
                    'MTC40F2046S1RC56BD1001', 'MTC40F2046S1RC56BD1002', 'MTC40F2046S1RC56BD100A', 'MTC40F2046S1RC56BD100B', 'MTC40F2046S1RC56BG1', 'MTC40F2046S1RC64BD1001', 'MTC40F2046S1RC64BD1003', 'MTC40F2046S1RC64BD100A', 'MTC40F2046S1RC64BD100B', 'MTC40F204WS1RC48BB1',
                    'N7K-AC-6.0KW', 'N7K-M224XP-23L', 'NCS1k-2KW-AC', 'NCS1k-FAN', 'NCS2K-100GS-CK-C', 'Non-Designated', 'Non-Designated', 'NXA-PAC-1100W-PI2', 'NXA-PAC-500W-PI', 'NXA-PAC-750W-PI', 'ONS-CFP2-WDM', 'OS2-8C-LC_ULL-LC_ULL-T-14B', 'P00918-B21', 'P21718-B21', 'P21718-L21',
                    'P787808GZ92A2GN008MB', 'P787808GZ92A2GN0100MB', 'P787808GZ92A2GN014MB', 'P787808GZ92A2GN016MB', 'P787808GZ92A2GN018MB', 'P787808GZ92A2GN022MB', 'P787808GZ92A2GN028MB', 'P787808GZ92A2GN035MB', 'P787808GZ92A2GN045MB', 'P9RN2', 'P9RN2', 'P9RN2', 'PF40281BX-Q501-S99',
                    'PF40561B1-D140-S99', 'PFR0612DHE', 'PS-2162-1L', 'PS-2322-1ADU-LF Rev.X6', 'PS-2422-1ADU Rev.X3', 'PS-2422-1ADU Rev.X4', 'PS-2422-1ADU Rev.X5', 'PS-2422-1ADU Rev.X6', 'PS7551BDVIHAF', 'PS7571BDVIHAF', 'PTX1000-FAN-S', 'PWR-C1-350WAC', 'PWR-MX480-2520-AC-BB',
                    'PWR-MX960-4100-AC-S', 'QAOC-10G4F1A01', 'QAOC-10G4F1A03', 'QAOC-10G4F1A03', 'QFX3500-FAN-AFO', 'QFX5200-32C-FANAFO', 'QSFP-4x10G-LR-S', 'R1K6A004L', 'R1P74', 'R1P74', 'R40W12BGE9-07T841', 'R40W12BGNL9-07AA43', 'R40W12BGNL9-07AA43', 'R40W12BS1NI9-07T022',
                    'R40W12BS1NI9-07T022', 'R40W12BS9NA9-07T022', 'R80W12BS3M9-57T111', 'RD-00010', 'RD-00122', 'RD-00125', 'RD-00127', 'RD-00138', 'RD-00144', 'RD-00145', 'RD-00219', 'RD-00348-001', 'RD-16GAH13R2B4L', 'RD-16GAS13R2B4L', 'RD-16GBSS13R2A4', 'RD-16GHY16R2B4L', 'RD-16MT16R2H4L',
                    'RD-2GBKT667E', 'RD-2GBWS667ER', 'RD-4GBAH13E2B2L', 'RD-4GBAH13R2F2I', 'RD-4GBAH16R2A4', 'RD-4GBCT1333E', 'RD-4GBCT1333E', 'RD-4GBHY1066ER2', 'RD-4GBHY1066ER2', 'RD-4GBHY1333ER2', 'RD-4GBHY1333ER2', 'RD-4GBHY13ER2B', 'RD-4GBHY13ER2B', 'RD-4GBHY13ER2B', 'RD-4GBHY13ER2C',
                    'RD-4GBHY13ER2C', 'RD-4GBHY13ER2C', 'RD-4GBHY13R2E2', 'RD-4GBHY13R2E2', 'RD-4GBHY13R2F2L', 'RD-4GBHY16R2A2', 'RD-4GBHY667ER2', 'RD-4GBHY667ER2', 'RD-4GBHY800ER2', 'RD-4GBHY800ER2', 'RD-4GBHY800ER2', 'RD-4GBHY800ER2B', 'RD-4GBKT667ER', 'RD-4GBMT1333ER', 'RD-4GBMT13ER2B',
                    'RD-4GBMT13R2B2L', 'RD-4GBMT16E1B4L', 'RD-4GBMT16R2A2', 'RD-4GBMT667ER', 'RD-4GBSS10ER2B', 'RD-4GBSS13ER2B', 'RD-4GBSS13ER2B', 'RD-4GBSS13ER2B', 'RD-4GBSS13ER2B', 'RD-4GBSS13ER2B', 'RD-4GBSS13R2C2', 'RD-4GBSS13R2C2', 'RD-4GBSS13R2C2', 'RD-4GBSS13R2C2', 'RD-4GBSS13R2C2',
                    'RD-4GBSS13R2D2L', 'RD-4GBSS13R2D2L', 'RD-4GBSS13R2F2L', 'RD-4GBSS667ER2', 'RD-4GBWS667ER', 'RD-4GBWS667ER', 'RD-4GBWS667ER', 'RD-4GBWS667ER', 'RD-8GBAH13R2D2L', 'RD-8GBAH16R2B2', 'RD-8GBAH16R2B4L', 'RD-8GBAM16E2B4L', 'RD-8GBAM16R2D4L', 'RD-8GBAS13R2B2L', 'RD-8GBAS16R2A2',
                    'RD-8GBAS16R2A4', 'RD-8GBAS16R2D4L', 'RD-8GBHY13R2D2L', 'RD-8GBHY13R2D2L', 'RD-8GBHY16E2B4L', 'RD-8GBHY16E2B4L', 'RD-8GBHY16R2A2', 'RD-8GBHY16R2B2', 'RD-8GBHY16R2B4L', 'RD-8GBHY18R2A4', 'RD-8GBHY18R2A4', 'RD-8GBMT13R2A2L', 'RD-8GBMT16E2B4L', 'RD-8GBMT16E2B4L', 'RD-8GBMT16R2D2L',
                    'RD-8GBSS1333ER2', 'RD-8GBSS1333ER2', 'RD-8GBSS133ER2', 'RD-8GBSS13R2A2L', 'RD-8GBSS13R2A2L', 'RD-8GBSS13R2B2L', 'RD-8GBSS13R2B2L', 'S5211G2NR-1U', 'S5530WGM2NR-LE-1T-AMZ-B', 'S5530WGM2NR-LE-1T-AMZ-B ECN C', 'S7063GM2NR-1T', 'S7063GM3NR-2T-B', 'S7063WGM2NR-1T',
                    'S7063WGM3NR-2T', 'S7067GM2NR-1T-B-SYN', 'S7067WGM2NR-1T', 'S7094GMR-LE-R0.4', 'S7094GMR-LE-SYN', 'S8010WGM2NR', 'S8010WGP2N-LE', 'S8212GM3NR', 'SFP-10GLR-31', 'SFP-GE-T',
                    'SM fanout 22mtr', 'SM fanout 35mtr', 'SM fanout 45mtr', 'SPQ-10E-SR-CDFF', 'SPQ-CE-LR-CDFB', 'SPQ-CE-LR-CDFL', 'SPTSBP3CLCAW', 'SPTSBP3CLCZA', 'SPTSBP4CLCAW', 'SPTSBP4CLCWA',
                    'SPTSHP3PMCAS', 'SPTSHP3PMCWA', 'SPTSHP3PMCWS', 'SPTSLP3SLCAS', 'SPTSLP3SLCAW', 'SPTSLP3SLCWS', 'SRX5800-PWR-4100-AC', 'SRX-CFP-100G-LR4', 'SRX-MIC-1X100G-CFP', 'SRX-QSFP-40G-LR4',
                    'SRX-QSFP-40G-SR4', 'SXP3101NVJ5', 'T192H', 'TG500', 'TJ1DY', 'TPD1XGKZRCAMZ21G', 'TPD1XGKZRCAMZ22G', 'TPD1XGKZRCAMZ23G', 'TPD1XGKZRCAMZ24G', 'TPD1XGKZRCAMZ25G',
                    'TPD1XGKZRCAMZ26G', 'TPD1XGKZRCAMZ27G', 'TPD1XGKZRCAMZ28G', 'TPD1XGKZRCAMZ29G', 'TPD1XGKZRCAMZ30G', 'TPD1XGKZRCAMZ31G', 'TPD1XGKZRCAMZ32G', 'TPD1XGKZRCAMZ33G', 'TPD1XGKZRCAMZ34G',
                    'TPD1XGKZRCAMZ35G', 'TPD1XGKZRCAMZ36G', 'TPD1XGKZRCAMZ37G', 'TPD1XGKZRCAMZ38G', 'TPD1XGKZRCAMZ39G', 'TPD1XGKZRCAMZ40G', 'TPD1XGKZRCAMZ41G', 'TPD1XGKZRCAMZ42G', 'TPD1XGKZRCAMZ43G',
                    'TPD1XGKZRCAMZ44G', 'TPD1XGKZRCAMZ45G', 'TPD1XGKZRCAMZ46G', 'TPD1XGKZRCAMZ47G', 'TPD1XGKZRCAMZ48G', 'TPD1XGKZRCAMZ49G', 'TPD1XGKZRCAMZ50G', 'TPD1XGKZRCAMZ51G', 'TPD1XGKZRCAMZ52G',
                    'TPD1XGKZRCAMZ53G', 'TPD1XGKZRCAMZ54G', 'TPD1XGKZRCAMZ55G', 'TPD1XGKZRCAMZ56G', 'TPD1XGKZRCAMZ57G', 'TPD1XGKZRCAMZ58G', 'TPD1XGKZRCAMZ60G', 'TP-VC13T01P8-NAM', 'TP-VC13T01P8-NW2',
                    'TP-VC13T02P7-NAM', 'TP-VC13T02P7-NW2', 'TR-FC13L-N00', 'TR-FC13R-NAM', 'TR-FC13T-NWB', 'TR-FC13T-NWD', 'TR-FC13X-NAM', 'TR-FC13X-NWB', 'TR-FC13X-NWD', 'TR-IQ13L-NAM', 'TR-IQ13L-NW1',
                    'TR-PX13L-N00', 'TR-ZC13H-H00', 'TR-ZC13T-NWB', 'TR-ZC13T-NWD', 'TYM-7063WGM2NR-1T-LE', 'TYM-7063WGM3NR-2T-LE', 'TYM-S7063GM2NR-1T-LE', 'TYM-S7063GM3NR-2T-LE', 'V80E12BS8NB5-07AA4',
                    'VG92561BX-Q010-S9H', 'VQ2830YP100', 'VQ2830YP150', 'VQ2830YP200', 'VQ2830YP250', 'X3R5M', 'X3R5M', 'XFP-10G-L-OC192-SR1', 'XW6VT', 'Y10KJ', 'Y10KJ', 'Y-CABLE', 'Z0W1-2000382096',
                    'Z-CAT6-12M-GN-UTP-AA-L-BBB', 'ZY372D3S4P13C9', 'NDAAFF0001', 'NDAAFF0002', 'NDAAFF0006', 'NDAAFF0007', 'NDAQGF0002', 'NDAQGF0005', 'NDYYJR-A302', 'NDYYJR-A304',
                    'MTA18ASF2G72PDZ-3G2E', 'MTA18ASF2G72PDZ-3G2E1', 'MTA36ASF4G72PZ-3G2E', 'MTA36ASF4G72PZ-3G2E7', 'MTA36ASF4G72PZ-3G2J', 'MTA36ASF4G72PZ-3G2J3', 'MTA36ASF8G72PZ-3G2E1',
                    '100-000000091', 'PS-2322-HADU-LF', 'AF800F00007', 'DAA-DPS-950AB-B', 'PS-2322-HADU-LF Rev.0A', 'PS-2302-1A1U-LF', 'DPST-3030AB C_02', 'PS-00301', 'MAZN860207F00-532-G', 'MAZN860207A00-504-G',
                    'LON-PS-4651-1HY1-LF', 'LON- PS-2122-8L', 'AF500B00025', 'DPS-650XB-2G S0F', 'MAZN860207700-065-G', 'MAZN860203U00-065-G', 'MAZN860206H00-289-G', 'PS-00295', 'DAA-DPS-650XB-2G', 'PS-00178',
                    '1HY9ZZZ028I', 'DAA-DPS-1200AB-13B02', 'AFC00B00028', 'DPST-3030AB C_00', 'DPS-950AB B S1F', 'EOLD-134HG-02-M51', 'EOLQ-131HG-O-02-51', 'MAZN800309E00-216-G', 'HMA81GR7CJR8N-WM', 'KAZ-AM16D426R19D8HAX',
                    'HYX-HMA82GR7CJR8NAWS', 'HMA84GR7CJR4N-XN', 'MAZNAM800308V00-026-G', 'HMA84GR7DJR4N-XN', 'HMAA8GL7CPR4N-WM', 'HMAA8GR7AJR4N-XN', '18ASF2G72PDZ-2G9E1', 'HMABAGR7A2R4N-XS', 'RD-00392', 'RD-00466',
                    'MAZN80030AF00-216-G', 'T-DP4CNT-NWB', 'KAZ-12D426R19D8HABX1', 'RD-00353-001', 'T-DP4CNT-NWD', 'RD-00407-001', 'TR-ZC13T-NWB', 'RD-8GBHY18R2C4', 'HYX-HMA84GR7CJR4NAWS', 'MAZN800309T00-026-G', 'RD-00527',
                    'SPTSBP3CLCZA', 'SAM-M393A1543CTDAWS', 'SPTSLP3SLCWS', 'MAZN80030A300-216-G', 'AM32D429R21D4HJX', 'MTA9ASF1G72PZ-2G9E1', '36ASF4G72PZ-3G2E7', 'RD-00411', '36ASF8G72PZ-3G2B2', '9ASF1G72PZ-2G9E1',
                    'MTA18ASF2G72PDZ-3G2E', 'HYA-CD8069504262901', 'MTA18ASF2G72PDZ-3G2E1', 'MTA36ASF4G72PZ-3G2E', 'TR-FC13T-NAZ', 'MTA36ASF4G72PZ-3G2E7', 'TR-ZC13T-NW5', 'FN-00286', 'MTA36ASF4G72PZ-3G2J', 'MTA18ASF2G72PDZ-2G9E',
                    'KAZ-M393A4K40CB2-CTD', 'MTA36ASF4G72PZ-3G2J3', 'MAZN80030A900-220-G', 'MAZN80030BS00-216-G', 'MTA36ASF8G72PZ-3G2E1', 'M386A8K40BM2-CVF', 'MTA18ASF2G72PDZ-2G9J1', 'M393A2K43BB3-CVF', 'M393A4K40BB3-CVF', 'MAZNTEMP-S000106095',
                    'M393A4K40EB3-CWE', 'TR-IQ13C-N00', 'HYX-HMA81GR7CJR8N-VK', 'M393A8G40BB4-CWE', 'MAZN800309S00-216-G', 'HYX-HMT42GR7DFR4-AWS', 'HYX-HMT42GR7BFR4A-PB', 'RD-00371-001', 'MAZN800309U00-220-G', 'MTA36ASF4G72PZ-2G9J1',
                    'MAZN490208R00-065-G', 'TR-FC13T-NWD', 'MTA72ASS8G72LZ-2G9J'



                ],

                'GERMANY': [
                    '1061700655-02'
                ],

                'KOREA': [
                    'HMA82GR7CJR8N-XN', 'HMA82GR7DJR8N-XN', 'HMAA8GR7AJR4N-WM', 'HMAA8GR7CJR4N-XN', 'M321R2GA3BB0-CQK', 'M393A2K43BB3-CWE',
                    'M393A4K40DB3-CWE', 'M393A8G40AB2-CWE', 'M393AAG40M32-CAE',
                    'M321R2GA3BB0-CWM', 'M321R2GA3BB6-CQK', 'M321R2GA3BB6-CQKD', 'M321R2GA3BB6-CQKE', 'M321R2GA3BB6-CQKM', 'M321R2GA3BB6-CQKV', 'M321R2GA3PB0-CWMC', 'M321R2GA3PB0-CWMJ', 'M321R2GA3PB0-CWMK', 'M321R2GA3PB0-CWMM', 'M321R2GA3PB0-CWMX', 'M321R2GA3PB1-CCPE', 'M321R2GA3PB1-CCPP',
                    'M321R2GA3PB1-CCPQ', 'M321R2GA3PB1-CCPY', 'M321R3GA3BB0-CQK', 'M321R3GA3BB0-CQKB', 'M321R3GA3BB0-CQKP', 'M321R3GA3BB0-CQKZ', 'M321R3GA3PB0-CWMC', 'M321R3GA3PB0-CWMK', 'M321R3GA3PB0-CWMM', 'M321R3GA3PB0-CWMX', 'M321R4GA0BB0-CQK', 'M321R4GA0BB0-CQKD', 'M321R4GA0BB0-CQKE',
                    'M321R4GA0BB0-CQKM', 'M321R4GA0BB0-CQKV', 'M321R4GA0BB0-CWM', 'M321R4GA0PB0-CWMC', 'M321R4GA0PB0-CWMK', 'M321R4GA0PB0-CWMM', 'M321R4GA0PB0-CWMX', 'M321R4GA3BB0-CQK', 'M321R4GA3BB0-CWM', 'M321R4GA3BB6-CQK', 'M321R4GA3BB6-CQKD', 'M321R4GA3BB6-CQKE', 'M321R4GA3BB6-CQKM',
                    'M321R4GA3BB6-CQKV', 'M321R4GA3PB0-CWMC', 'M321R4GA3PB0-CWMJ', 'M321R4GA3PB0-CWMK', 'M321R4GA3PB0-CWMM', 'M321R4GA3PB0-CWMX', 'M321R4GA3PB1-CCPE', 'M321R4GA3PB1-CCPP', 'M321R4GA3PB1-CCPQ', 'M321R4GA3PB1-CCPY', 'M321R8GA0BB0-CQK', 'M321R8GA0BB0-CQKD', 'M321R8GA0BB0-CQKE',
                    'M321R8GA0BB0-CQKM', 'M321R8GA0BB0-CQKR', 'M321R8GA0BB0-CQKV', 'M321R8GA0BB0-CQKZ', 'M321R8GA0BB0-CWM', 'M321R8GA0PB0-CWM', 'M321R8GA0PB0-CWMC', 'M321R8GA0PB0-CWMJ', 'M321R8GA0PB0-CWMK', 'M321R8GA0PB0-CWMM', 'M321R8GA0PB0-CWMX', 'M321R8GA0PB1-CCPE', 'M321R8GA0PB1-CCPP',
                    'M321R8GA0PB1-CCPQ', 'M321R8GA0PB1-CCPY', 'M321RAGA0B20-CWK', 'M321RAGA0B20-CWKB', 'M321RAGA0B20-CWKP', 'M321RAGA0B20-CWKZ', 'M321RBGA0B40-CWK', 'M321RBGA0B40-CWKB', 'M321RBGA0B40-CWKP', 'M321RBGA0B40-CWKZ', 'M321RCGA0B60-CWK', 'M321RYGA0BB0-CQKB', 'M321RYGA0BB0-CQKP',
                    'M321RYGA0BB0-CQKZ', 'M322R8GA0BB0-CQK', 'M322RAGA0B20-CWK', 'M322RBGA0B40-CWK', 'M322RCGA0B60-CWK', 'M329R8GA0BB0-CQK', 'M386A4G40EM2-CTD', 'HYX-HMA84GR7J4NVKAWS', 'HMAA8GR7CJR4N-XN', 'HYX-HMA82GR7AFR8-AWS', 'ATR3GAGRW19', 'ATR5XAGRL04', 'M393A2K43BB3-CWE', 'M393A4K40DB3-CWE', 'M393A8G40AB2-CWE', 'ATR5CAGRL00', 'ATR3XAGML12',
                    'ATR3GAGML36', 'ATR3XAGM503', 'ATR38AGRP02', 'ATR38AGMW33', 'ATR38AGMW51'


                ],

                'MALAYSIA': [
                    'AFCT-89SFDZ-AZ4', 'FCBG110SD1C03', 'FCBN410QB1C03', 'FCBN410QB1C03', 'FCBN425QE1C05', 'FTL410QD4C', 'FTL410QE1C', 'FTL4C1QE1L', 'FTL4C1QE2L',
                    'FTL4C1QL1L', 'FTL4C1QL2L', 'FTL4E1QE1C', 'FTLC1151RDPL2', 'FTLC1151RDPL2-BZ', 'FTLC1152RGPL2', 'FTLC1183RDNL', 'FTLF8519P3BNL',
                    'FCBG110SD1C01', 'FCBG110SD1C01', 'FCBG110SD1C01', 'FCBG110SD1C02', 'FCBG110SD1C02', 'FCBG110SD1CC5', 'FCBN425QB1C03', 'FCBN425QE1C03', 'SFP-1GE-SX', 'SFP-1GE-T', 'SRX-SFP-1GE-LH', 'SRX-SFP-1GE-SX', 'SRX-SFP-1GE-SX',
                    'CP-00381-002', 'CP-00493', 'CP-00454', 'AJSRFA8RA00'


                ],

                'THAILAND': [
                    'CS8000-32X-DC-11 (R04)', 'SPTSBP3CLCAZ', 'SPTSBP4CLCAZ', 'CD8067303173000', 'CD8067303535900', 'CD8067303561400',
                    'CD8067303589801', 'CD8067304123001', 'CD8069504228201', 'CD8069504262801', 'CD8069504262901', 'CD8069504425100',
                    'CD8069504425200', 'CD8069504448600', 'CM8066002402400', 'CM8066002685800', 'CM8066002685800', 'EA100WDM02A24-M',
                    'EA100WDM02A24-R', 'EA100WDM02AFB01', 'EA100WDM10A24-M', 'EA100WDM10A24-R', 'EA100WDM10AFB01', 'SPTSBP3CLCSA',
                    'SPTSBP3CLCSW', 'SPTSBP3CLCWS', 'SPTSBP3PCCWA002', 'SPTSBP3PCCWA003', 'SPTSBP3PCCWS002', 'SPTSBP3PCCWS006',
                    'SPTSBP4CLCSA', 'SPTSBP4CLCWS', 'TRQ5E20ENF-LF000', 'SPTSBP3CLCSA',
                    'CS8000-32X-DC-11', 'AD610EHDK42GM', 'AT80615007089AA', 'BX80623I32100', 'CD8067302577901', 'CD8067302577901-951612', 'CD8067303133605', 'CD8067303172900', 'CD8067303562100', 'CD8067303567200',
                    'CD8068904586205', 'CD8068904586205-99A9', 'CD8068904589003', 'CD8068904704602', 'CD8069503969304', 'CD8069503969402', 'CD8069503969500', 'CD8069504144901', 'CD8069504185501', 'CD8069504185701',
                    'CD8069504189801', 'CD8069504189801', 'CD8069504190002', 'CD8069504262701', 'CM8062100856218', 'CM8062100856401', 'CM8062101082713', 'CM8062107184801', 'CM8062301061600', 'CM8063501288100',
                    'CM8063501374901', 'CM8063501375000', 'CM8063501375101', 'CM8063501375101', 'CM8063501376200', 'CM8063501452503', 'CM8063501521101', 'CM8063701160503', 'CM8064401438110', 'CM8064401613101',
                    'CM8064401675902', 'CM8064601467204', 'CM8066002032701', 'CM8066002034302', 'CM8066002396203', 'CM8066003197800', 'TRS7081FN21PA000', 'TRS7081FN22PA000', 'TRS7081FN23PA000', 'TRS7081FN24PA000',
                    'TRS7081FN25PA000', 'TRS7081FN26PA000', 'TRS7081FN27PA000', 'TRS7081FN28PA000', 'TRS7081FN29PA000', 'TRS7081FN30PA000', 'TRS7081FN31PA000', 'TRS7081FN32PA000', 'TRS7081FN33PA000', 'TRS7081FN34PA000',
                    'TRS7081FN35PA000', 'TRS7081FN36PA000', 'TRS7081FN37PA000', 'TRS7081FN38PA000', 'TRS7081FN39PA000', 'TRS7081FN40PA000', 'TRS7081FN41PA000', 'TRS7081FN42PA000', 'TRS7081FN43PA000', 'TRS7081FN44PA000',
                    'TRS7081FN45PA000', 'TRS7081FN46PA000', 'TRS7081FN47PA000', 'TRS7081FN48PA000', 'TRS7081FN49PA000', 'TRS7081FN50PA000', 'TRS7081FN51PA000', 'TRS7081FN52PA000', 'TRS7081FN53PA000', 'TRS7081FN54PA000',
                    'TRS7081FN55PA000', 'TRS7081FN56PA000', 'TRS7081FN57PA000', 'TRS7081FN58PA000', 'TRS7081FN59PA000', 'TRS7081FN60PA000'


                ],

                'TAIWAN': [
                    '7760-32X-A-AC-F', '7762-32X-A-12V-F-21', '7762-32X-A-AC-F', '7762-32X-A-AC-F-D-HVDC', '7772-32X-A-AC-F', '7772-32X-A-AC-F-G-HVDC',
                    '7772-32X-A-AC-F-I-HVDC', 'K2C-A1', 'K2C-AB', 'K2C-ABM-T2', 'K2T-25x16410010161SS-FHS', 'K2T-NS1',
                    'K2T-QB', 'K2T-QB-1TPM', 'K2T-QB-T2', 'K2T-QB-T2A', 'K2T-QB-TP1', 'K2V4-N',
                    'K2V4-N2', 'K2X-A-32G-T2', 'K2X-A-32G-T2A', 'K2X-N', 'NT-00138-002', 'ANP-K2T-QB',
                    'ANP-K2C-AB', 'K2T-25X16410010161SS-FHS', 'K2X-N R02', 'MAZN210-000185-001', 'K2C-NS-T2', 'K2B-L1',
                    'K2C-NS1', 'NT-00120', 'K2C-25x1600004160X-LPB', 'K2T-25x16410010161TX-FHS', 'K2C-NS', 'K2B-N1',
                    'K2C-ABM-T2', 'PX2-2-T-1 01 R0C V3A', 'PX2-2-T-1', 'PX2-2-T-1A', '31S2RMB00L0', '383-19830-3001',
                    '383-19830-3003', '383-19830-3004', '383-19830-3005', 'CVR-QSFP-SFP10G', 'Dummy-001', 'F0TAO7632001A', 'F0TAO7632008A',
                    'FPR9K-NM-4x40G', 'FPR9K-NM-8X10G', 'GLC-LH-SM', 'GLC-SX-MM', 'GLC-TE', 'NT16GA72D8PBX3P-IX', 'ONS-SC-OSC-ULH',
                    'ONS-SE-155-1510', 'QSFP-40G-CSR-S', 'QSFP-40G-LR4-S', 'QSFP-40G-SR-BD', 'YM-2651YBR', '7772-32X-A-12V-F-21', 'F0OAO8632001A', 'F0TAO6632004A',
                    '7772-32X-A-12V-F-21', '130-04056-3130_A0', '1395T2380106', '31S2RMB00S0', '31S2RMB00T0', '31S2VMB0090', '31S4LMB0040', 'F0OAO8632001A', 'F0TAO6632004A', 'NT16GA72D8PBX3P-HR',
                    'NT16GA72D8PFX3K-JR', 'NT16GA72D8PFX3P-JR', 'NT32GA72D4NBX3P-HR', 'NT32GA72D4NBX3P-IX', 'NT32GA72D4NFX3K-JR', 'NT32GA72D4NFX3P-JR', 'NT4GC72B8PB', 'NT5AD1024M8C3-JR', 'NT5AD1024M8F3-JR',
                    'NT5AD256M16E4-JR', 'NT5AD512M16F4-JR', 'NT8GA72D89BX3P-IX', 'NT8GA72D89FX3K-JR', 'NT8GA72D89FX3P-JR', 'NT8GC72C4NB3NJ', 'NT8GC72C4NG0NL-CG',
                    '124-397760-001', '124-700031-001', 'AS7762-32X', '7762-32X-A-AC-F-HVDC', '7762-32X-A-12V-F-21-D', '313-000084-001', 'HYX-HMA84GR7CJR4N-VK', 'KAZ-AM32D426R19D4MEX', 'SAM-M393A4K40BCTDAWS', 'FAN-1U-1x1C-F1', 'UNKNOWN_DONGLE'



                ],
                'US': [
                    'AFCT-91DRPHZ-AZ2', 'S2600IP4', 'XQX4002',
                    '500658-B21', '501535-001', '501536-001', '725585-B21', '805349-B21', '815097-B21', '815098-B21', '835955-B21', '860649-L21', '867958-B21',
                    'AFCT-91DRPHZ-AZ4', 'H8SCM-F', 'H8SCM-F', 'H8SCM-F', 'MBD-H8SCM-F-AM041', 'MEM-DR340L-CL02-ER13', 'MEM-DR340L-CL02-ER13', 'MEM-DR340L-HL02-ER13', 'MEM-DR340L-HL04-ER13', 'MEM-DR340L-HL05-ER13',
                    'MEM-DR340L-HL05-ER13', 'MEM-DR340L-HL05-ER13', 'MEM-DR340L-HL06-ER13', 'MEM-DR340L-HL08-ER13', 'MEM-DR340L-SL04-ER13', 'MEM-DR340L-SL06-ER13', 'MEM-DR340L-SL06-ER13', 'MEM-DR340L-SL06-ER13',
                    'MEM-DR380L-HL03-ER13', 'MEM-DR380L-HL06-ER16', 'MEM-DR380L-SL03-ER13', 'P41682-001'

                ],
                'PHILIPPINES': [
                    '9CRV0412P5J220', '9CRV0412P5J221', '9GA0412P3K26', '9GA0612P1J601', '9GA0612P1J711', '9HV0812P1G6051'

                ],

                'JAPAN': [
                    'FC9542TRC2', 'HA74L-0001-0160'
                ]
            };

            ////////////////////////////////COO CODE//////////////////////////////////////////////////////////



            Object.entries(cooToMPNs).forEach(([country, mpns]) => {
                if (mpns.includes(mpnCellText)) {
                    rowDataArray[9] = country; // Set CoO if the MPN matches any of the MPNs for the country
                }
            });

            const mpnKey = rowDataArray[3]; // Assuming MPN is at index 3 in rowDataArray
            if (mpnKey.includes('UNKNOWN')) {
                unknownData.push(rowDataArray);
                rowDataArray[1] = '<< Need Re-ID!!'; // Set Brand to "<< Need Re-ID!!" for UNKNOWN lines
            } else {
                // Check if the MPN has already been added to knownData
                const existingRow = knownData.find(row => row[3] === mpnKey);
                if (existingRow) {
                    // Update the quantity in the current row
                    rowDataArray[10] += 1; // Increment quantity in current row
                } else {
                    // Add a new row with the MPN and quantity 1
                    knownData.push(rowDataArray);
                }
                // Increment count for the MPN
                if (mpnCountMap.has(mpnKey)) {
                    mpnCountMap.set(mpnKey, mpnCountMap.get(mpnKey) + 1);
                } else {
                    mpnCountMap.set(mpnKey, 1); // Initialize count for new MPN
                }
            }
            console.log("MPN Count Map:", mpnCountMap);

        });

        // Combine unknown and known data, with unknown data first
        extractedData.push(...unknownData, ...knownData);

        // Update Qty column with count of duplicates
        extractedData.forEach(rowDataArray => {
            const mpnKey = rowDataArray[3];
            if (!mpnKey.includes('UNKNOWN')) {
                rowDataArray.splice(10, 1, mpnCountMap.get(mpnKey));
            }
        });

        const headers = [
            'Asset ID', 'Brand', 'IPN', 'MPN', 'Description', 'Origin HS Code', 'Destination HS Code',
            'ECCN', 'Lic/ Lic Exc', 'CoO', 'Qty', 'Net Wt', 'Actual Unit Cost', 'Total Value'
        ]; // Updated with provided column headers

        const csvContent = [
            headers.join(','),
            ...extractedData.map(row => row.map(cell => cell === null ? '' : `"${cell.toString().replace(/"/g, '""')}"`).join(','))
        ].join('\n');

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);

        const filename = `CI TEMPLATE DATA TO PASTE.csv`;

        link.setAttribute('href', url);
        link.setAttribute('download', filename);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    // Event listener to trigger the CSV download on button click
    document.getElementById('createCIDownloadButton').addEventListener('click', copySpecificColumnsToCSV3);



})();

// ==/UserScript Continuation== 
// moorpatx@ is gonna be updating and adding more as time goes by for L4 FEE work
//Things to do:
//1. User Preference Storage
//2. Export/Import Settings
//3. Keyboard Shortcuts
//4. Enhanced Search Functionality
//5. Data Analytics Dashboard
//6. Batch Operations History
//7. Enhanced Error Handling
//8. Auto-Refresh
//9. Custom Views/Layouts
//10. make this mobile responsive *why not?*