/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 */
define(['N/ui/serverWidget', 'N/search', 'N/record', 'N/redirect', 'N/log', 'N/runtime'],
    /**
     * @param {serverWidget} serverWidget
     * @param {search} search
     * @param {record} record
     * @param {redirect} redirect
     * @param {log} log
     * @param {runtime} runtime
     */
    function (serverWidget, search, record, redirect, log, runtime) {

        /**
         * Handles GET and POST requests to the Suitelet
         * @param {Object} context - NetSuite context object containing request/response
         */
        function onRequest(context) {
            if (context.request.method === 'GET') {
                handleGet(context);
            } else {
                handlePost(context);
            }
        }

        /**
         * Handles GET requests - displays the variance table
         * @param {Object} context
         */
        function handleGet(context) {
            var request = context.request;
            var response = context.response;

            log.debug('GET Request', 'Parameters: ' + JSON.stringify(request.parameters));

            // Create NetSuite form
            var form = serverWidget.createForm({
                title: 'Purchase Order vs Vendor Bill Rate Variance Review'
            });

            try {
                // Build and add HTML content
                var htmlContent = buildPageHTML(request.parameters);

                var htmlField = form.addField({
                    id: 'custpage_html_content',
                    type: serverWidget.FieldType.INLINEHTML,
                    label: 'Content'
                });
                htmlField.defaultValue = htmlContent;

            } catch (e) {
                log.error('Error Building Page', e.toString());
                var errorField = form.addField({
                    id: 'custpage_error',
                    type: serverWidget.FieldType.INLINEHTML,
                    label: 'Error'
                });
                errorField.defaultValue = '<div style="color: red; padding: 20px;">Error: ' + escapeHtml(e.toString()) + '</div>';
            }

            response.writePage(form);
        }

        /**
         * Handles POST requests - processes selected variance reviews
         * @param {Object} context
         */
        function handlePost(context) {
            var request = context.request;
            var selectedVariances = request.parameters.selected_variances;
            var batchIndex = parseInt(request.parameters.batch_index || '0');
            var locationFilter = request.parameters.location_filter || 'all';

            // Preserve threshold parameters if they exist
            var serviceThreshold = request.parameters.service_threshold;
            var kitchenThreshold = request.parameters.kitchen_threshold;
            var appliancesThreshold = request.parameters.appliances_threshold;

            log.audit('POST Request - Variance Review', {
                totalSelected: selectedVariances ? selectedVariances.split(',').length : 0,
                batchIndex: batchIndex,
                locationFilter: locationFilter
            });

            if (!selectedVariances) {
                var redirectParams = {
                    error: 'No variances selected',
                    location_filter: locationFilter
                };

                // Preserve thresholds in redirect
                if (serviceThreshold) redirectParams.service_threshold = serviceThreshold;
                if (kitchenThreshold) redirectParams.kitchen_threshold = kitchenThreshold;
                if (appliancesThreshold) redirectParams.appliances_threshold = appliancesThreshold;

                redirect.toSuitelet({
                    scriptId: runtime.getCurrentScript().id,
                    deploymentId: runtime.getCurrentScript().deploymentId,
                    parameters: redirectParams
                });
                return;
            }

            var allUpdates = selectedVariances.split(',');
            var batchSize = 10; // Process 10 records per batch
            var startIndex = batchIndex * batchSize;
            var endIndex = Math.min(startIndex + batchSize, allUpdates.length);
            var currentBatch = allUpdates.slice(startIndex, endIndex);

            // Track cumulative results
            var previousSuccessCount = parseInt(request.parameters.success_count || '0');
            var previousErrorCount = parseInt(request.parameters.error_count || '0');
            var previousErrors = request.parameters.previous_errors ? JSON.parse(request.parameters.previous_errors) : [];
            var previousUpdated = request.parameters.previous_updated ? JSON.parse(request.parameters.previous_updated) : [];

            var successCount = 0;
            var errorCount = 0;
            var errors = [];
            var updatedRecords = [];

            // Process current batch
            currentBatch.forEach(function (update) {
                try {
                    var parts = update.split('|');
                    var poId = parts[0];
                    var poLineKey = parts[1];
                    var poNumber = parts[2];
                    var itemName = parts[3];

                    log.debug('Processing Review', {
                        poId: poId,
                        poLineKey: poLineKey,
                        poNumber: poNumber,
                        batchIndex: batchIndex,
                        recordIndex: startIndex + currentBatch.indexOf(update)
                    });

                    // Mark PO line as reviewed
                    markPOLineReviewed(poId, poLineKey);

                    successCount++;
                    updatedRecords.push({
                        poNumber: poNumber,
                        itemName: itemName
                    });

                } catch (e) {
                    errorCount++;

                    var errorMessage = e.message;
                    if (errorMessage.indexOf('closed period') !== -1) {
                        errorMessage = 'Period is closed - cannot modify transaction';
                    } else if (errorMessage.indexOf('No lines found') !== -1) {
                        errorMessage = 'Line not found on Purchase Order';
                    } else if (errorMessage.indexOf('governance') !== -1 || errorMessage.indexOf('Usage Limit') !== -1) {
                        errorMessage = 'Script usage limit exceeded';
                    }

                    errors.push({
                        poId: parts[0],
                        poNumber: parts[2] || 'Unknown',
                        itemName: parts[3] || 'Unknown',
                        error: errorMessage
                    });

                    log.error('PO Update Failed - Continuing to Next', {
                        poId: parts[0],
                        poNumber: parts[2],
                        itemName: parts[3],
                        error: errorMessage,
                        fullError: e.message
                    });
                }
            });

            // Combine with previous results
            var totalSuccessCount = previousSuccessCount + successCount;
            var totalErrorCount = previousErrorCount + errorCount;
            var allErrors = previousErrors.concat(errors);
            var allUpdated = previousUpdated.concat(updatedRecords);

            // Check if there are more batches to process
            if (endIndex < allUpdates.length) {
                log.audit('Batch Complete - Continuing', {
                    batchIndex: batchIndex,
                    processed: endIndex,
                    total: allUpdates.length,
                    remaining: allUpdates.length - endIndex
                });

                var continueParams = {
                    selected_variances: selectedVariances,
                    batch_index: batchIndex + 1,
                    success_count: totalSuccessCount,
                    error_count: totalErrorCount,
                    previous_errors: JSON.stringify(allErrors),
                    previous_updated: JSON.stringify(allUpdated),
                    processing: 'true',
                    location_filter: locationFilter
                };

                // Preserve thresholds in redirect
                if (serviceThreshold) continueParams.service_threshold = serviceThreshold;
                if (kitchenThreshold) continueParams.kitchen_threshold = kitchenThreshold;
                if (appliancesThreshold) continueParams.appliances_threshold = appliancesThreshold;

                redirect.toSuitelet({
                    scriptId: runtime.getCurrentScript().id,
                    deploymentId: runtime.getCurrentScript().deploymentId,
                    parameters: continueParams
                });
            } else {
                log.audit('All Batches Complete', {
                    totalBatches: batchIndex + 1,
                    totalSuccess: totalSuccessCount,
                    totalErrors: totalErrorCount
                });

                var completeParams = {
                    updateSuccess: 'true',
                    successCount: totalSuccessCount,
                    errorCount: totalErrorCount,
                    errors: totalErrorCount > 0 ? JSON.stringify(allErrors) : null,
                    updatedRecords: JSON.stringify(allUpdated),
                    location_filter: locationFilter
                };

                // Preserve thresholds in redirect
                if (serviceThreshold) completeParams.service_threshold = serviceThreshold;
                if (kitchenThreshold) completeParams.kitchen_threshold = kitchenThreshold;
                if (appliancesThreshold) completeParams.appliances_threshold = appliancesThreshold;

                redirect.toSuitelet({
                    scriptId: runtime.getCurrentScript().id,
                    deploymentId: runtime.getCurrentScript().deploymentId,
                    parameters: completeParams
                });
            }
        }

        /**
         * Marks a PO line as reviewed by setting the Rate Variance Reviewed checkbox
         * @param {string} poId - Purchase Order internal ID
         * @param {string} poLineKey - PO Line unique key
         */
        function markPOLineReviewed(poId, poLineKey) {
            log.debug('Marking PO Line Reviewed', {
                poId: poId,
                poLineKey: poLineKey
            });

            var poRecord = record.load({
                type: record.Type.PURCHASE_ORDER,
                id: poId,
                isDynamic: false
            });

            var lineCount = poRecord.getLineCount({ sublistId: 'item' });
            var lineFound = false;

            for (var i = 0; i < lineCount; i++) {
                var currentLineKey = poRecord.getSublistValue({
                    sublistId: 'item',
                    fieldId: 'lineuniquekey',
                    line: i
                });

                if (currentLineKey && currentLineKey.toString() === poLineKey.toString()) {
                    poRecord.setSublistValue({
                        sublistId: 'item',
                        fieldId: 'custcol_rate_variance_reviewed',
                        line: i,
                        value: true
                    });

                    lineFound = true;
                    log.debug('Line Found and Marked', {
                        line: i,
                        lineKey: currentLineKey
                    });
                    break;
                }
            }

            if (!lineFound) {
                throw new Error('No lines found with Line Key ' + poLineKey + ' on PO ' + poId);
            }

            var savedId = poRecord.save({
                enableSourcing: false,
                ignoreMandatoryFields: true
            });

            log.audit('PO Line Marked as Reviewed', {
                poId: savedId,
                lineKey: poLineKey
            });

            return savedId;
        }

        /**
         * Gets script parameters for variance thresholds
         * @returns {Object} Variance threshold percentages by location type
         */
        function getVarianceThresholds() {
            var script = runtime.getCurrentScript();

            var servicePercent = parseFloat(script.getParameter({ name: 'custscript_service_variance_percent' }) || 2);
            var kitchensPercent = parseFloat(script.getParameter({ name: 'custscript_kitchens_variance_percent' }) || 2);
            var appliancesPercent = parseFloat(script.getParameter({ name: 'custscript_appliances_variance_percent' }) || 2);

            log.debug('Variance Thresholds', {
                service: servicePercent + '%',
                kitchens: kitchensPercent + '%',
                appliances: appliancesPercent + '%'
            });

            return {
                service: servicePercent,
                kitchens: kitchensPercent,
                appliances: appliancesPercent
            };
        }

        /**
         * Gets the variance threshold for a specific location
         * @param {string} locationId - Location internal ID
         * @param {Object} thresholds - Variance thresholds object
         * @returns {number} Variance threshold percentage
         */
        function getThresholdForLocation(locationId, thresholds) {
            // Service location: 113
            if (locationId === '113') {
                return thresholds.service;
            }
            // Kitchen Works location: 17
            else if (locationId === '17') {
                return thresholds.kitchens;
            }
            // Appliances: all others
            else {
                return thresholds.appliances;
            }
        }

        /**
         * Builds the main page HTML content
         * @param {Object} params - URL parameters
         * @returns {string} HTML content
         */
        function buildPageHTML(params) {
            var html = '<style>' + getStyles() + '</style>';
            html += '<div class="container">';

            var locationFilter = params.location_filter || 'all';

            // Get thresholds (check if overridden in URL params first)
            var thresholds = getVarianceThresholds();
            var serviceThreshold = params.service_threshold ? parseFloat(params.service_threshold) : thresholds.service;
            var kitchenThreshold = params.kitchen_threshold ? parseFloat(params.kitchen_threshold) : thresholds.kitchens;
            var appliancesThreshold = params.appliances_threshold ? parseFloat(params.appliances_threshold) : thresholds.appliances;

            // Override thresholds object if params provided
            if (params.service_threshold || params.kitchen_threshold || params.appliances_threshold) {
                thresholds = {
                    service: serviceThreshold,
                    kitchens: kitchenThreshold,
                    appliances: appliancesThreshold
                };
            }

            // Show processing message if in batch mode
            if (params.processing === 'true') {
                var successCount = parseInt(params.success_count || '0');
                var errorCount = parseInt(params.error_count || '0');
                var totalProcessed = successCount + errorCount;

                html += '<div class="processing-message">';
                html += '<strong>⏳ Processing Updates...</strong><br />';
                html += 'Records processed: ' + totalProcessed;
                html += '<br />Successful: ' + successCount;
                if (errorCount > 0) {
                    html += '<br />Failed: ' + errorCount;
                }
                html += '<br /><br />';
                html += '<div class="spinner"></div>';
                html += '<br />Please wait while the remaining records are updated...';
                html += '</div>';

                html += '<form id="continueForm" method="POST" style="display:none;">';
                html += '<input type="hidden" name="selected_variances" value="' + escapeHtml(params.selected_variances || '') + '" />';
                html += '<input type="hidden" name="batch_index" value="' + escapeHtml(params.batch_index || '0') + '" />';
                html += '<input type="hidden" name="success_count" value="' + escapeHtml(params.success_count || '0') + '" />';
                html += '<input type="hidden" name="error_count" value="' + escapeHtml(params.error_count || '0') + '" />';
                html += '<input type="hidden" name="previous_errors" value="' + escapeHtml(params.previous_errors || '[]') + '" />';
                html += '<input type="hidden" name="previous_updated" value="' + escapeHtml(params.previous_updated || '[]') + '" />';
                html += '<input type="hidden" name="location_filter" value="' + escapeHtml(locationFilter) + '" />';
                html += '</form>';
                html += '<script>setTimeout(function() { document.getElementById("continueForm").submit(); }, 1000);</script>';
                html += '</div>';
                return html;
            }

            // Show success message
            if (params.updateSuccess === 'true') {
                html += buildUpdateSuccessMessage(params);
            }

            // Show error message (but not when just changing filter)
            if (params.error && !params.location_filter) {
                html += '<div class="error-message">';
                html += '<strong>✗ Error</strong><br />';
                html += escapeHtml(params.error);
                html += '</div>';
            }

            // Add filter section with location and thresholds
            html += '<div class="filter-section">';
            html += '<form method="GET" id="filterForm">';

            // Location Filter
            html += '<div style="margin-bottom: 15px;">';
            html += '<label for="location_filter" style="font-weight: bold; margin-right: 10px;">Filter by Location:</label>';
            html += '<select id="location_filter" name="location_filter" style="padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px;">';
            html += '<option value="all"' + (locationFilter === 'all' ? ' selected' : '') + '>All Locations</option>';
            html += '<option value="service"' + (locationFilter === 'service' ? ' selected' : '') + '>Service</option>';
            html += '<option value="kitchen"' + (locationFilter === 'kitchen' ? ' selected' : '') + '>Kitchen Works</option>';
            html += '<option value="appliances"' + (locationFilter === 'appliances' ? ' selected' : '') + '>Appliances</option>';
            html += '</select>';
            html += '</div>';

            // Variance Thresholds
            html += '<div class="threshold-container">';
            html += '<label style="font-weight: bold; margin-right: 15px;">Variance Thresholds (%):</label>';
            html += '<div class="threshold-inputs">';

            html += '<div class="threshold-group">';
            html += '<label for="service_threshold">Service:</label>';
            html += '<input type="number" id="service_threshold" name="service_threshold" value="' + serviceThreshold + '" step="0.1" min="0" style="width: 60px; padding: 6px; border: 1px solid #ccc; border-radius: 4px; text-align: right;" />';
            html += '<span>%</span>';
            html += '</div>';

            html += '<div class="threshold-group">';
            html += '<label for="kitchen_threshold">Kitchen Works:</label>';
            html += '<input type="number" id="kitchen_threshold" name="kitchen_threshold" value="' + kitchenThreshold + '" step="0.1" min="0" style="width: 60px; padding: 6px; border: 1px solid #ccc; border-radius: 4px; text-align: right;" />';
            html += '<span>%</span>';
            html += '</div>';

            html += '<div class="threshold-group">';
            html += '<label for="appliances_threshold">Appliances:</label>';
            html += '<input type="number" id="appliances_threshold" name="appliances_threshold" value="' + appliancesThreshold + '" step="0.1" min="0" style="width: 60px; padding: 6px; border: 1px solid #ccc; border-radius: 4px; text-align: right;" />';
            html += '<span>%</span>';
            html += '</div>';

            html += '<button type="submit" class="apply-threshold-button">Apply</button>';
            html += '</div>';
            html += '</div>';

            html += '</form>';
            html += '</div>';

            // Get variance data with current thresholds
            var variancePairs = getVariancePairsWithThresholds(locationFilter, thresholds);

            if (variancePairs.length === 0) {
                html += '<div class="info-message">';
                html += '<strong>ℹ No Variances Found</strong><br />';
                html += 'No unreviewed variances found for the selected location filter that meet the variance threshold.';
                html += '</div>';
            } else {
                html += '<div class="summary-info">';
                html += '<strong>Total Variances Found:</strong> ' + variancePairs.length + ' line(s)';
                var filterText = locationFilter === 'service' ? ' (Service)' :
                    locationFilter === 'kitchen' ? ' (Kitchen Works)' :
                        locationFilter === 'appliances' ? ' (Appliances)' : '';
                if (filterText) {
                    html += ' <span style="color: #666;">' + filterText + '</span>';
                }
                html += '</div>';
                html += buildVarianceTable(variancePairs, locationFilter);
            }

            html += '</div>';
            html += '<script>' + getJavaScript() + '</script>';

            return html;
        }

        /**
         * Builds success message after update
         * @param {Object} params - URL parameters
         * @returns {string} HTML content
         */
        function buildUpdateSuccessMessage(params) {
            var html = '<div class="success-message">';
            html += '<strong>✓ Review Process Complete</strong><br />';
            html += params.successCount + ' line(s) marked as reviewed successfully.';

            if (params.errorCount > 0) {
                html += '<br /><br />';
                html += '<span style="color: #d32f2f; font-weight: bold;">⚠ ' + params.errorCount + ' line(s) could not be updated:</span>';

                if (params.errors) {
                    try {
                        var errors = JSON.parse(params.errors);
                        html += '<div style="margin-top: 10px; max-height: 300px; overflow-y: auto; border: 1px solid #f5c6cb; border-radius: 4px; padding: 10px; background: #fff;">';
                        html += '<table style="width: 100%; border-collapse: collapse;">';
                        html += '<thead style="position: sticky; top: 0; background: #f8d7da;">';
                        html += '<tr style="border-bottom: 2px solid #f5c6cb;">';
                        html += '<th style="text-align: left; padding: 8px;">PO #</th>';
                        html += '<th style="text-align: left; padding: 8px;">Item</th>';
                        html += '<th style="text-align: left; padding: 8px;">Reason</th>';
                        html += '</tr>';
                        html += '</thead>';
                        html += '<tbody>';
                        errors.forEach(function (err) {
                            html += '<tr style="border-bottom: 1px solid #f5c6cb;">';
                            html += '<td style="padding: 8px;"><a href="/app/accounting/transactions/purchord.nl?id=' + err.poId + '" target="_blank">' + escapeHtml(err.poNumber) + '</a></td>';
                            html += '<td style="padding: 8px;">' + escapeHtml(err.itemName) + '</td>';
                            html += '<td style="padding: 8px; color: #721c24;">' + escapeHtml(err.error) + '</td>';
                            html += '</tr>';
                        });
                        html += '</tbody>';
                        html += '</table>';
                        html += '</div>';
                    } catch (e) {
                        log.error('Error Parsing Error Details', e);
                    }
                }
            }

            // Show updated records summary
            if (params.updatedRecords) {
                try {
                    var updated = JSON.parse(params.updatedRecords);
                    if (updated.length > 0) {
                        html += '<div style="margin-top: 15px; padding: 10px; background: #f5f5f5; border-radius: 4px;">';
                        html += '<strong>✓ Successfully Marked as Reviewed:</strong>';
                        html += '<div style="max-height: 200px; overflow-y: auto; margin-top: 5px;">';
                        html += '<ul style="margin: 0; padding-left: 20px;">';
                        updated.forEach(function (rec) {
                            html += '<li>' + escapeHtml(rec.poNumber) + ' - ' + escapeHtml(rec.itemName) + '</li>';
                        });
                        html += '</ul>';
                        html += '</div>';
                        html += '</div>';
                    }
                } catch (e) {
                    log.error('Error Parsing Updated Records', e);
                }
            }

            html += '</div>';
            return html;
        }

        /**
         * Builds the variance table HTML
         * @param {Array} variancePairs - Array of variance pair objects
         * @param {string} locationFilter - Current location filter
         * @returns {string} HTML table content
         */
        function buildVarianceTable(variancePairs, locationFilter) {
            var html = '<form id="varianceForm" method="POST">';
            html += '<input type="hidden" name="location_filter" value="' + escapeHtml(locationFilter || 'all') + '" />';

            html += '<table class="variance-table">';
            html += '<thead>';
            html += '<tr>';
            html += '<th class="checkbox-col"><input type="checkbox" id="selectAll" title="Select/Deselect All" /></th>';
            html += '<th>VB Date</th>';
            html += '<th>VB #</th>';
            html += '<th>Location</th>';
            html += '<th>Vendor</th>';
            html += '<th class="variance-cell">Variance %</th>';
            html += '<th class="variance-cell">Variance $</th>';
            html += '<th>Item</th>';
            html += '<th>PO #</th>';
            html += '<th class="rate-cell">VB Rate</th>';
            html += '<th class="rate-cell">PO Rate</th>';
            html += '</tr>';
            html += '</thead>';
            html += '<tbody>';

            variancePairs.forEach(function (pair) {
                var variance = pair.vb_rate - pair.po_rate;
                var variancePercent = pair.po_rate !== 0 ? (variance / pair.po_rate * 100) : 0;

                // Negative variance = good (paid less), Positive variance = bad (paid more)
                var varianceClass = '';
                if (variancePercent < -0.1) {
                    varianceClass = 'variance-good'; // Green for negative (paid less)
                } else if (variancePercent > 0.1) {
                    varianceClass = 'variance-bad'; // Red for positive (paid more)
                }

                // Create unique value for checkbox: PO_ID|PO_LINE_KEY|PO_NUMBER|ITEM_NAME
                var checkboxValue = pair.po_id + '|' +
                    pair.po_line_key + '|' +
                    pair.po_number + '|' +
                    pair.item_name;

                html += '<tr>';
                html += '<td class="checkbox-col"><input type="checkbox" class="variance-checkbox" value="' + escapeHtml(checkboxValue) + '" /></td>';
                html += '<td>' + formatDate(pair.vb_date) + '</td>';
                html += '<td><a href="/app/accounting/transactions/vendbill.nl?id=' + pair.vb_id + '" target="_blank">' + escapeHtml(pair.vb_number) + '</a></td>';
                html += '<td>' + escapeHtml(pair.location_name || '') + '</td>';
                html += '<td>' + escapeHtml(pair.vendor_name || '') + '</td>';
                html += '<td class="variance-cell ' + varianceClass + '">' + variancePercent.toFixed(1) + '%</td>';
                html += '<td class="variance-cell ' + varianceClass + '">$' + variance.toFixed(2) + '</td>';
                html += '<td>' + escapeHtml(pair.item_name) + '</td>';
                html += '<td><a href="/app/accounting/transactions/purchord.nl?id=' + pair.po_id + '" target="_blank">' + escapeHtml(pair.po_number) + '</a></td>';
                html += '<td class="rate-cell">$' + pair.vb_rate.toFixed(2) + '</td>';
                html += '<td class="rate-cell">$' + pair.po_rate.toFixed(2) + '</td>';
                html += '</tr>';
            });

            html += '</tbody>';
            html += '</table>';

            html += '<div class="button-container">';
            html += '<button type="button" class="submit-button" onclick="submitVariances()">Mark Selected as Reviewed</button>';
            html += '</div>';

            html += '</form>';

            return html;
        }

        /**
         * Formats a date string
         * @param {string} dateStr - Date string
         * @returns {string} Formatted date
         */
        function formatDate(dateStr) {
            if (!dateStr) return '';
            try {
                var date = new Date(dateStr);
                var month = ('0' + (date.getMonth() + 1)).slice(-2);
                var day = ('0' + date.getDate()).slice(-2);
                var year = date.getFullYear();
                return month + '/' + day + '/' + year;
            } catch (e) {
                return dateStr;
            }
        }

        /**
        * Gets variance pairs with custom thresholds
        * @param {string} locationFilter - Location filter (all/service/kitchen/appliances)
        * @param {Object} thresholds - Custom variance thresholds
        * @returns {Array} Array of variance pair objects
        */
        function getVariancePairsWithThresholds(locationFilter, thresholds) {
            var rawResults = searchPOVBVariances(locationFilter);
            var poLineGroups = groupByPOLine(rawResults);
            var variancePairs = createVariancePairs(poLineGroups, thresholds);

            // Sort by VB Date descending (newest first)
            variancePairs.sort(function (a, b) {
                return new Date(b.vb_date) - new Date(a.vb_date);
            });

            log.debug('Variance Pairs Created', 'Total pairs: ' + variancePairs.length);

            return variancePairs;
        }

        /**
  * Searches for PO/VB rate variances
  * @param {string} locationFilter - Location filter
  * @returns {Array} Raw search results
  */
        function searchPOVBVariances(locationFilter) {
            var filters = [
                ['type', 'anyof', 'PurchOrd'],
                'AND',
                ['mainline', 'is', 'F'],
                'AND',
                ['billingtransaction.quantity', 'greaterthan', '0'],
                'AND',
                ['quantity', 'greaterthan', '0'],
                'AND',
                ['formulanumeric: NVL({rate},0)-NVL({billingtransaction.amount}/{billingtransaction.quantity},0)', 'notequalto', '0'],
                'AND',
                ['custcol_rate_variance_reviewed', 'is', 'F'],
                'AND',
                ['billingtransaction.trandate', 'onorafter', '8/1/2025']
            ];

            // Add location filter
            if (locationFilter === 'service') {
                filters.push('AND', ['location', 'anyof', '113']);
            } else if (locationFilter === 'kitchen') {
                filters.push('AND', ['location', 'anyof', '17']);
            } else if (locationFilter === 'appliances') {
                filters.push('AND', ['location', 'noneof', '113', '17']);
            }

            var varianceSearch = search.create({
                type: search.Type.TRANSACTION,
                filters: filters,
                columns: [
                    search.createColumn({ name: 'internalid', label: 'PO ID' }),
                    search.createColumn({ name: 'tranid', label: 'PO Number' }),
                    search.createColumn({ name: 'trandate', label: 'PO Date' }),
                    search.createColumn({ name: 'entity', label: 'Vendor ID' }),
                    search.createColumn({ name: 'altname', join: 'vendor', label: 'Vendor Display Name' }), // ADD THIS LINE
                    search.createColumn({ name: 'entityid', join: 'vendor', label: 'Vendor Name' }), // ADD THIS LINE
                    search.createColumn({ name: 'location', label: 'Location ID' }),
                    search.createColumn({ name: 'name', join: 'location', label: 'Location Name' }),
                    search.createColumn({ name: 'lineuniquekey', label: 'PO Line Key' }),
                    search.createColumn({ name: 'item', label: 'Item ID' }),
                    search.createColumn({ name: 'displayname', join: 'item', label: 'Item Name' }),
                    search.createColumn({ name: 'rate', label: 'PO Rate' }),
                    search.createColumn({ name: 'quantity', label: 'PO Quantity' }),
                    // Vendor Bill columns
                    search.createColumn({ name: 'internalid', join: 'billingtransaction', label: 'VB ID' }),
                    search.createColumn({ name: 'tranid', join: 'billingtransaction', label: 'VB Number' }),
                    search.createColumn({ name: 'trandate', join: 'billingtransaction', label: 'VB Date' }),
                    search.createColumn({ name: 'lineuniquekey', join: 'billingtransaction', label: 'VB Line Key' }),
                    search.createColumn({ name: 'quantity', join: 'billingtransaction', label: 'VB Quantity' }),
                    search.createColumn({ name: 'rate', join: 'billingtransaction', label: 'VB Rate' }),
                    search.createColumn({ name: 'entity', join: 'billingtransaction', label: 'VB Vendor ID' })
                ]
            });

            var results = [];

            varianceSearch.run().each(function (result) {
                var itemName = result.getText({ name: 'item' }) || result.getValue({ name: 'displayname', join: 'item' }) || '';

                // Get vendor display name from the search results directly (no lookup needed)
                var vendorName = result.getValue({ name: 'altname', join: 'vendor' }) ||
                    result.getValue({ name: 'entityid', join: 'vendor' }) ||
                    result.getText({ name: 'entity' }) ||
                    'Unknown Vendor';

                results.push({
                    po_id: result.getValue({ name: 'internalid' }),
                    po_number: result.getValue({ name: 'tranid' }),
                    po_date: result.getValue({ name: 'trandate' }),
                    vendor_id: result.getValue({ name: 'entity' }),
                    vendor_name: vendorName, // UPDATED THIS
                    location_id: result.getValue({ name: 'location' }),
                    location_name: result.getValue({ name: 'name', join: 'location' }),
                    po_line_key: result.getValue({ name: 'lineuniquekey' }),
                    item_id: result.getValue({ name: 'item' }),
                    item_name: itemName,
                    po_rate: result.getValue({ name: 'rate' }),
                    po_quantity: result.getValue({ name: 'quantity' }),
                    vb_id: result.getValue({ name: 'internalid', join: 'billingtransaction' }),
                    vb_number: result.getValue({ name: 'tranid', join: 'billingtransaction' }),
                    vb_date: result.getValue({ name: 'trandate', join: 'billingtransaction' }),
                    vb_line_key: result.getValue({ name: 'lineuniquekey', join: 'billingtransaction' }),
                    vb_quantity: result.getValue({ name: 'quantity', join: 'billingtransaction' }),
                    vb_rate: result.getValue({ name: 'rate', join: 'billingtransaction' })
                });
                return true;
            });

            log.debug('Search Results', 'Total rows: ' + results.length);
            return results;
        }

        /**
         * Groups raw query results by PO Line Key
         * @param {Array} rawResults - Raw search results
         * @returns {Object} Grouped results by PO line
         */
        function groupByPOLine(rawResults) {
            var groups = {};

            rawResults.forEach(function (row) {
                var poLineKey = row.po_line_key;

                if (!groups[poLineKey]) {
                    groups[poLineKey] = {
                        poInfo: {
                            po_id: row.po_id,
                            po_number: row.po_number,
                            po_date: row.po_date,
                            vendor_id: row.vendor_id,
                            vendor_name: row.vendor_name,
                            location_id: row.location_id,
                            location_name: row.location_name,
                            po_line_key: row.po_line_key,
                            item_id: row.item_id,
                            item_name: row.item_name,
                            po_rate: parseFloat(row.po_rate),
                            po_quantity: parseFloat(row.po_quantity)
                        },
                        vendorBills: []
                    };
                }

                // Add Vendor Bill if not already added
                var vbExists = groups[poLineKey].vendorBills.some(function (vb) {
                    return vb.vb_line_key === row.vb_line_key;
                });
                if (!vbExists) {
                    groups[poLineKey].vendorBills.push({
                        vb_id: row.vb_id,
                        vb_number: row.vb_number,
                        vb_date: row.vb_date,
                        vb_line_key: row.vb_line_key,
                        vb_quantity: parseFloat(row.vb_quantity),
                        vb_rate: parseFloat(row.vb_rate)
                    });
                }
            });

            log.debug('Grouped by PO Line', 'Total PO lines with variances: ' + Object.keys(groups).length);
            return groups;
        }

        /**
         * Creates variance pairs by matching PO to oldest VB
         * @param {Object} poLineGroups - Grouped results
         * @param {Object} thresholds - Variance thresholds by location type
         * @returns {Array} Array of variance pair objects
         */
        function createVariancePairs(poLineGroups, thresholds) {
            var pairs = [];

            Object.keys(poLineGroups).forEach(function (poLineKey) {
                var group = poLineGroups[poLineKey];

                // Sort VBs by date (oldest first)
                group.vendorBills.sort(function (a, b) {
                    return new Date(a.vb_date) - new Date(b.vb_date);
                });

                // Match PO line to oldest VB
                if (group.vendorBills.length > 0) {
                    var vb = group.vendorBills[0]; // Oldest VB
                    var variance = vb.vb_rate - group.poInfo.po_rate;
                    var variancePercent = group.poInfo.po_rate !== 0 ? (variance / group.poInfo.po_rate * 100) : 0;

                    // Get threshold for this location
                    var threshold = getThresholdForLocation(group.poInfo.location_id, thresholds);

                    // Only include if variance >= $0.01 AND meets percentage threshold
                    // Show if abs(variancePercent) >= threshold
                    if (Math.abs(variance) >= 0.01 && Math.abs(variancePercent) >= threshold) {
                        pairs.push({
                            po_id: group.poInfo.po_id,
                            po_number: group.poInfo.po_number,
                            po_date: group.poInfo.po_date,
                            vendor_id: group.poInfo.vendor_id,
                            vendor_name: group.poInfo.vendor_name,
                            location_id: group.poInfo.location_id,
                            location_name: group.poInfo.location_name,
                            po_line_key: group.poInfo.po_line_key,
                            item_id: group.poInfo.item_id,
                            item_name: group.poInfo.item_name,
                            po_rate: group.poInfo.po_rate,
                            po_quantity: group.poInfo.po_quantity,
                            vb_id: vb.vb_id,
                            vb_number: vb.vb_number,
                            vb_date: vb.vb_date,
                            vb_line_key: vb.vb_line_key,
                            vb_quantity: vb.vb_quantity,
                            vb_rate: vb.vb_rate
                        });
                    }
                }
            });

            return pairs;
        }

        /**
         * Escapes HTML special characters
         * @param {string} text - Text to escape
         * @returns {string} Escaped text
         */
        function escapeHtml(text) {
            if (!text) return '';
            var map = {
                '&': '&amp;',
                '<': '&lt;',
                '>': '&gt;',
                '"': '&quot;',
                "'": '&#039;'
            };
            return text.toString().replace(/[&<>"']/g, function (m) { return map[m]; });
        }

        /**
         * Returns CSS styles for the page
         * @returns {string} CSS content
         */
        function getStyles() {
            return `
                * {
                    box-sizing: border-box;
                }
                
                body {
                    font-family: Arial, sans-serif;
                    margin: 0;
                    padding: 0;
                    background: #f5f5f5;
                }
                
                .container {
                    max-width: 1800px;
                    margin: 0 auto;
                    padding: 20px;
                }
                
                .filter-section {
                    background: white;
                    padding: 15px 20px;
                    border-radius: 8px;
                    margin-bottom: 20px;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                }
                
                .threshold-container {
                    padding-top: 10px;
                    border-top: 1px solid #e0e0e0;
                    margin-top: 10px;
                }
                
                .threshold-inputs {
                    display: flex;
                    align-items: center;
                    gap: 20px;
                    flex-wrap: wrap;
                    margin-top: 8px;
                }
                
                .threshold-group {
                    display: flex;
                    align-items: center;
                    gap: 6px;
                }
                
                .threshold-group label {
                    font-size: 13px;
                    color: #666;
                }
                
                .threshold-group span {
                    font-size: 13px;
                    color: #666;
                }
                
                .apply-threshold-button {
                    background: #34a853;
                    color: white;
                    border: none;
                    padding: 6px 16px;
                    font-size: 13px;
                    font-weight: 600;
                    border-radius: 4px;
                    cursor: pointer;
                    transition: background 0.2s;
                }
                
                .apply-threshold-button:hover {
                    background: #2d9048;
                }
                
                .apply-threshold-button:active {
                    background: #27803d;
                }
                
                .summary-info {
                    background: #e3f2fd;
                    padding: 15px;
                    border-radius: 8px;
                    margin-bottom: 20px;
                    border-left: 4px solid #1a73e8;
                    font-size: 16px;
                }
                
                .processing-message {
                    background: #fff3cd;
                    border: 1px solid #ffc107;
                    color: #856404;
                    padding: 20px;
                    border-radius: 8px;
                    margin-bottom: 20px;
                    text-align: center;
                    font-size: 16px;
                }
                
                .spinner {
                    border: 4px solid #f3f3f3;
                    border-top: 4px solid #1a73e8;
                    border-radius: 50%;
                    width: 40px;
                    height: 40px;
                    animation: spin 1s linear infinite;
                    margin: 10px auto;
                }
                
                @keyframes spin {
                    0% { transform: rotate(0deg); }
                    100% { transform: rotate(360deg); }
                }
                
                .success-message {
                    background: #d4edda;
                    border: 1px solid #c3e6cb;
                    color: #155724;
                    padding: 15px;
                    border-radius: 8px;
                    margin-bottom: 20px;
                }
                
                .error-message {
                    background: #f8d7da;
                    border: 1px solid #f5c6cb;
                    color: #721c24;
                    padding: 15px;
                    border-radius: 8px;
                    margin-bottom: 20px;
                }
                
                .info-message {
                    background: #d1ecf1;
                    border: 1px solid #bee5eb;
                    color: #0c5460;
                    padding: 15px;
                    border-radius: 8px;
                    margin-bottom: 20px;
                }
                
                .variance-table {
                    width: 100%;
                    border-collapse: collapse;
                    background: white;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                    border-radius: 8px;
                    overflow: hidden;
                    font-family: Arial, sans-serif;
                    font-size: 14px;
                }
                
                .variance-table thead {
                    background: #1a73e8;
                    color: white;
                }
                
                .variance-table th,
                .variance-table td {
                    padding: 10px;
                    text-align: left;
                    border-bottom: 1px solid #e0e0e0;
                    font-family: Arial, sans-serif;
                    font-size: 14px;
                }
                
                .variance-table th {
                    font-weight: 600;
                    font-size: 13px;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                }
                
                .variance-table tbody tr:hover {
                    background: #f5f5f5;
                }
                
                .variance-table a {
                    color: #1a73e8;
                    text-decoration: none;
                    font-family: Arial, sans-serif;
                    font-size: 14px;
                }
                
                .variance-table a:hover {
                    text-decoration: underline;
                }
                
                .checkbox-col {
                    width: 40px;
                    text-align: center;
                }
                
                .rate-cell {
                    text-align: right;
                    font-family: Arial, sans-serif;
                    font-size: 14px;
                }
                
                .variance-cell {
                    text-align: right;
                    font-family: Arial, sans-serif;
                    font-weight: bold;
                    font-size: 14px;
                }
                
                .variance-good {
                    color: #2e7d32;
                    background-color: #e8f5e9;
                }
                
                .variance-bad {
                    color: #d32f2f;
                    background-color: #ffebee;
                }
                
                .button-container {
                    margin-top: 20px;
                    text-align: center;
                    padding: 20px;
                    background: white;
                    border-radius: 8px;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                }
                
                .submit-button {
                    background: #1a73e8;
                    color: white;
                    border: none;
                    padding: 12px 30px;
                    font-size: 16px;
                    font-weight: 600;
                    border-radius: 4px;
                    cursor: pointer;
                    transition: background 0.2s;
                }
                
                .submit-button:hover {
                    background: #1557b0;
                }
                
                .submit-button:active {
                    background: #0d47a1;
                }
                
                #selectAll {
                    cursor: pointer;
                    width: 18px;
                    height: 18px;
                }
                
                .variance-checkbox {
                    cursor: pointer;
                    width: 18px;
                    height: 18px;
                }
            `;
        }

        /**
        * Returns JavaScript for the page
        * @returns {string} JavaScript content
        */
        function getJavaScript() {
            return `
                document.addEventListener('DOMContentLoaded', function() {
                    var selectAll = document.getElementById('selectAll');
                    if (selectAll) {
                        selectAll.addEventListener('change', function() {
                            var checkboxes = document.querySelectorAll('.variance-checkbox');
                            checkboxes.forEach(function(cb) {
                                cb.checked = selectAll.checked;
                            });
                        });
                    }
                    
                    // Handle location filter change - preserve threshold values
                    var locationFilter = document.getElementById('location_filter');
                    if (locationFilter) {
                        locationFilter.addEventListener('change', function() {
                            // Form will auto-submit with all threshold inputs included
                            document.getElementById('filterForm').submit();
                        });
                    }
                });
                
                function submitVariances() {
                    var checkboxes = document.querySelectorAll('.variance-checkbox:checked');
                    
                    if (checkboxes.length === 0) {
                        alert('Please select at least one variance to mark as reviewed.');
                        return;
                    }
                    
                    var confirmMessage = 'Mark ' + checkboxes.length + ' PO line(s) as reviewed?\\n\\n';
                    confirmMessage += 'This will set the "Rate Variance Reviewed" checkbox on the selected PO lines.\\n';
                    confirmMessage += 'This does NOT change any rates.\\n\\n';
                    confirmMessage += 'Continue?';
                    
                    if (!confirm(confirmMessage)) {
                        return;
                    }
                    
                    var selected = [];
                    checkboxes.forEach(function(cb) {
                        selected.push(cb.value);
                    });
                    
                    var form = document.getElementById('varianceForm') || document.querySelector('form');
                    
                    if (!form) {
                        alert('Error: Could not find form element. Please refresh and try again.');
                        return;
                    }
                    
                    // Add selected variances
                    var input = document.createElement('input');
                    input.type = 'hidden';
                    input.name = 'selected_variances';
                    input.value = selected.join(',');
                    form.appendChild(input);
                    
                    // Add threshold values from the filter form
                    var serviceThreshold = document.getElementById('service_threshold');
                    var kitchenThreshold = document.getElementById('kitchen_threshold');
                    var appliancesThreshold = document.getElementById('appliances_threshold');
                    
                    if (serviceThreshold && serviceThreshold.value) {
                        var serviceInput = document.createElement('input');
                        serviceInput.type = 'hidden';
                        serviceInput.name = 'service_threshold';
                        serviceInput.value = serviceThreshold.value;
                        form.appendChild(serviceInput);
                    }
                    
                    if (kitchenThreshold && kitchenThreshold.value) {
                        var kitchenInput = document.createElement('input');
                        kitchenInput.type = 'hidden';
                        kitchenInput.name = 'kitchen_threshold';
                        kitchenInput.value = kitchenThreshold.value;
                        form.appendChild(kitchenInput);
                    }
                    
                    if (appliancesThreshold && appliancesThreshold.value) {
                        var appliancesInput = document.createElement('input');
                        appliancesInput.type = 'hidden';
                        appliancesInput.name = 'appliances_threshold';
                        appliancesInput.value = appliancesThreshold.value;
                        form.appendChild(appliancesInput);
                    }
                    
                    var submitButton = document.querySelector('.submit-button');
                    if (submitButton) {
                        submitButton.disabled = true;
                        submitButton.textContent = 'Marking ' + checkboxes.length + ' record(s)...';
                    }
                    
                    form.submit();
                }
            `;
        }

        return {
            onRequest: onRequest
        };
    });