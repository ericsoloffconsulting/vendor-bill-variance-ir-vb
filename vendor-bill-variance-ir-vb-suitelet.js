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
                title: 'Item Receipt vs Vendor Bill Rate Variance'
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
       * Handles POST requests - processes selected variance updates
       * @param {Object} context
       */
        function handlePost(context) {
            var request = context.request;

            // Check if this is a closed period adjustment request
            if (request.parameters.action === 'closed_period_adjustment') {
                handleClosedPeriodAdjustment(context);
                return;
            }

            var request = context.request;
            var selectedVariances = request.parameters.selected_variances;
            var batchIndex = parseInt(request.parameters.batch_index || '0');

            log.audit('POST Request - Variance Update', {
                totalSelected: selectedVariances ? selectedVariances.split(',').length : 0,
                batchIndex: batchIndex
            });

            if (!selectedVariances) {
                redirect.toSuitelet({
                    scriptId: runtime.getCurrentScript().id,
                    deploymentId: runtime.getCurrentScript().deploymentId,
                    parameters: { error: 'No variances selected' }
                });
                return;
            }

            var allUpdates = selectedVariances.split(',');
            var batchSize = 5; // Process 5 records per batch to stay under governance limits
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
                    var irId = parts[0];
                    var poLineId = parts[1];
                    var newRate = parts[2];
                    var irNumber = parts[3];
                    var itemName = parts[4];
                    var itemId = parts[5];

                    log.debug('Processing Update', {
                        irId: irId,
                        itemId: itemId,
                        newRate: newRate,
                        batchIndex: batchIndex,
                        recordIndex: startIndex + currentBatch.indexOf(update)
                    });

                    // Update the IR line rate
                    updateItemReceiptLineRate(irId, poLineId, itemId, newRate);

                    successCount++;
                    updatedRecords.push({
                        irNumber: irNumber,
                        itemName: itemName,
                        newRate: newRate
                    });

                } catch (e) {
                    errorCount++;

                    // Simplify error message for common cases
                    var errorMessage = e.message;
                    if (errorMessage.indexOf('closed period') !== -1) {
                        errorMessage = 'Period is closed - cannot modify GL impact';
                    } else if (errorMessage.indexOf('No lines found') !== -1) {
                        errorMessage = 'Item not found on Item Receipt';
                    } else if (errorMessage.indexOf('governance') !== -1 || errorMessage.indexOf('Usage Limit') !== -1) {
                        errorMessage = 'Script usage limit exceeded';
                    }

                    errors.push({
                        irId: parts[0],
                        irNumber: parts[3] || 'Unknown',
                        itemName: parts[4] || 'Unknown',
                        error: errorMessage
                    });

                    log.error('IR Update Failed - Continuing to Next', {
                        irId: parts[0],
                        irNumber: parts[3],
                        itemName: parts[4],
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
                // Continue to next batch
                log.audit('Batch Complete - Continuing', {
                    batchIndex: batchIndex,
                    processed: endIndex,
                    total: allUpdates.length,
                    remaining: allUpdates.length - endIndex,
                    currentBatchSuccess: successCount,
                    currentBatchErrors: errorCount
                });

                redirect.toSuitelet({
                    scriptId: runtime.getCurrentScript().id,
                    deploymentId: runtime.getCurrentScript().deploymentId,
                    parameters: {
                        selected_variances: selectedVariances,
                        batch_index: batchIndex + 1,
                        success_count: totalSuccessCount,
                        error_count: totalErrorCount,
                        previous_errors: JSON.stringify(allErrors),
                        previous_updated: JSON.stringify(allUpdated),
                        processing: 'true'
                    }
                });
            } else {
                // All batches complete - show final results
                log.audit('All Batches Complete', {
                    totalBatches: batchIndex + 1,
                    totalSuccess: totalSuccessCount,
                    totalErrors: totalErrorCount
                });

                redirect.toSuitelet({
                    scriptId: runtime.getCurrentScript().id,
                    deploymentId: runtime.getCurrentScript().deploymentId,
                    parameters: {
                        updateSuccess: 'true',
                        successCount: totalSuccessCount,
                        errorCount: totalErrorCount,
                        errors: totalErrorCount > 0 ? JSON.stringify(allErrors) : null,
                        updatedRecords: JSON.stringify(allUpdated)
                    }
                });
            }
        }

        // Add this new function after handlePost:
        /**
         * Handles closed period adjustment request
         * @param {Object} context
         */
        function handleClosedPeriodAdjustment(context) {
            var request = context.request;
            var vbId = request.parameters.vb_id;
            var itemId = request.parameters.item_id;
            var vbRate = parseFloat(request.parameters.vb_rate);
            var irRate = parseFloat(request.parameters.ir_rate);
            var vbNumber = request.parameters.vb_number;
            var itemName = request.parameters.item_name;

            try {
                log.audit('Closed Period Adjustment Started', {
                    vbId: vbId,
                    itemId: itemId,
                    vbRate: vbRate,
                    irRate: irRate
                });

                // Step 1: Load vendor bill and get original total
                var vbRecord = record.load({
                    type: record.Type.VENDOR_BILL,
                    id: vbId,
                    isDynamic: false
                });

                var originalTotal = vbRecord.getValue({ fieldId: 'total' });
                var department = null;

                // Step 2: Find and update the item line
                var itemLineCount = vbRecord.getLineCount({ sublistId: 'item' });
                var itemLineFound = false;

                for (var i = 0; i < itemLineCount; i++) {
                    var lineItem = vbRecord.getSublistValue({
                        sublistId: 'item',
                        fieldId: 'item',
                        line: i
                    });

                    if (lineItem && lineItem.toString() === itemId.toString()) {
                        var oldRate = vbRecord.getSublistValue({
                            sublistId: 'item',
                            fieldId: 'rate',
                            line: i
                        });

                        // Get department from this line
                        department = vbRecord.getSublistValue({
                            sublistId: 'item',
                            fieldId: 'department',
                            line: i
                        });

                        // Update rate to match IR
                        vbRecord.setSublistValue({
                            sublistId: 'item',
                            fieldId: 'rate',
                            line: i,
                            value: irRate
                        });

                        itemLineFound = true;
                        log.debug('Item Line Updated', {
                            line: i,
                            oldRate: oldRate,
                            newRate: irRate,
                            department: department
                        });
                        break;
                    }
                }

                if (!itemLineFound) {
                    throw new Error('Item not found on Vendor Bill');
                }

                // Step 3: Calculate adjustment amount (difference to offset)
                var adjustmentAmount = vbRate - irRate;

                // Step 4: Add expense line to offset the difference
                var expenseLineCount = vbRecord.getLineCount({ sublistId: 'expense' });

                vbRecord.insertLine({
                    sublistId: 'expense',
                    line: expenseLineCount
                });

                vbRecord.setSublistValue({
                    sublistId: 'expense',
                    fieldId: 'account',
                    line: expenseLineCount,
                    value: '112' // Accrued Purchases
                });

                vbRecord.setSublistValue({
                    sublistId: 'expense',
                    fieldId: 'amount',
                    line: expenseLineCount,
                    value: adjustmentAmount
                });

                var memo = 'Closed Period Adj: Item ' + itemName + ' (ID: ' + itemId + ') - ' +
                    'Orig VB Rate: $' + vbRate.toFixed(2) + ', ' +
                    'IR Rate: $' + irRate.toFixed(2) + ', ' +
                    'Diff: $' + adjustmentAmount.toFixed(2);

                vbRecord.setSublistValue({
                    sublistId: 'expense',
                    fieldId: 'memo',
                    line: expenseLineCount,
                    value: memo
                });

                // Step 5: Validate total hasn't changed
                var newTotal = vbRecord.getValue({ fieldId: 'total' });

                if (Math.abs(newTotal - originalTotal) > 0.01) {
                    throw new Error('VB total changed from $' + originalTotal.toFixed(2) + ' to $' + newTotal.toFixed(2) + ' - adjustment cancelled');
                }

                // Save the vendor bill
                var savedVbId = vbRecord.save({
                    enableSourcing: false,
                    ignoreMandatoryFields: true
                });

                log.audit('Vendor Bill Updated', {
                    vbId: savedVbId,
                    adjustmentAmount: adjustmentAmount
                });

                // Step 6: Create offsetting Journal Entry
                var jeRecord = record.create({
                    type: record.Type.JOURNAL_ENTRY,
                    isDynamic: false
                });

                jeRecord.setValue({
                    fieldId: 'trandate',
                    value: new Date()
                });

                jeRecord.setValue({
                    fieldId: 'memo',
                    value: 'Closed Period Adjustment for VB ' + vbNumber + ' - Item: ' + itemName
                });

                // Determine debit/credit based on adjustment amount
                if (adjustmentAmount > 0) {
                    // VB expense was positive, so CREDIT Accrued Purchases, DEBIT COGS

                    // Line 1: CREDIT Accrued Purchases (112)
                    jeRecord.setSublistValue({
                        sublistId: 'line',
                        fieldId: 'account',
                        line: 0,
                        value: '112'
                    });
                    jeRecord.setSublistValue({
                        sublistId: 'line',
                        fieldId: 'credit',
                        line: 0,
                        value: adjustmentAmount
                    });
                    jeRecord.setSublistValue({
                        sublistId: 'line',
                        fieldId: 'memo',
                        line: 0,
                        value: 'Offset accrued purchases - VB ' + vbNumber
                    });

                    // Line 2: DEBIT COGS (353)
                    jeRecord.setSublistValue({
                        sublistId: 'line',
                        fieldId: 'account',
                        line: 1,
                        value: '353'
                    });
                    jeRecord.setSublistValue({
                        sublistId: 'line',
                        fieldId: 'debit',
                        line: 1,
                        value: adjustmentAmount
                    });
                    jeRecord.setSublistValue({
                        sublistId: 'line',
                        fieldId: 'memo',
                        line: 1,
                        value: 'COGS adjustment for ' + itemName
                    });

                    // Set department on COGS line
                    var cogsDept = department === '13' ? '13' : (department === '10' ? '10' : '107');
                    jeRecord.setSublistValue({
                        sublistId: 'line',
                        fieldId: 'department',
                        line: 1,
                        value: cogsDept
                    });

                } else {
                    // VB expense was negative, so DEBIT Accrued Purchases, CREDIT COGS
                    var absAmount = Math.abs(adjustmentAmount);

                    // Line 1: DEBIT Accrued Purchases (112)
                    jeRecord.setSublistValue({
                        sublistId: 'line',
                        fieldId: 'account',
                        line: 0,
                        value: '112'
                    });
                    jeRecord.setSublistValue({
                        sublistId: 'line',
                        fieldId: 'debit',
                        line: 0,
                        value: absAmount
                    });
                    jeRecord.setSublistValue({
                        sublistId: 'line',
                        fieldId: 'memo',
                        line: 0,
                        value: 'Offset accrued purchases - VB ' + vbNumber
                    });

                    // Line 2: CREDIT COGS (353)
                    jeRecord.setSublistValue({
                        sublistId: 'line',
                        fieldId: 'account',
                        line: 1,
                        value: '353'
                    });
                    jeRecord.setSublistValue({
                        sublistId: 'line',
                        fieldId: 'credit',
                        line: 1,
                        value: absAmount
                    });
                    jeRecord.setSublistValue({
                        sublistId: 'line',
                        fieldId: 'memo',
                        line: 1,
                        value: 'COGS adjustment for ' + itemName
                    });

                    // Set department on COGS line
                    var cogsDept = department === '13' ? '13' : (department === '10' ? '10' : '107');
                    jeRecord.setSublistValue({
                        sublistId: 'line',
                        fieldId: 'department',
                        line: 1,
                        value: cogsDept
                    });
                }

                var jeId = jeRecord.save();
                var jeNumber = record.load({
                    type: record.Type.JOURNAL_ENTRY,
                    id: jeId
                }).getValue({ fieldId: 'tranid' });

                log.audit('Journal Entry Created', {
                    jeId: jeId,
                    jeNumber: jeNumber
                });

                // Redirect with success message
                redirect.toSuitelet({
                    scriptId: runtime.getCurrentScript().id,
                    deploymentId: runtime.getCurrentScript().deploymentId,
                    parameters: {
                        adjustmentSuccess: 'true',
                        vbNumber: vbNumber,
                        jeNumber: jeNumber,
                        itemName: itemName,
                        adjustmentAmount: adjustmentAmount.toFixed(2)
                    }
                });

            } catch (e) {
                log.error('Closed Period Adjustment Failed', e);

                redirect.toSuitelet({
                    scriptId: runtime.getCurrentScript().id,
                    deploymentId: runtime.getCurrentScript().deploymentId,
                    parameters: {
                        error: 'Adjustment failed: ' + e.message,
                        vbNumber: vbNumber
                    }
                });
            }
        }

        /**
         * Updates a specific line on an Item Receipt with new rate
         * @param {string} irId - Item Receipt internal ID
         * @param {string} poLineId - PO Line unique ID (may be undefined)
         * @param {string} itemId - Item internal ID
         * @param {number} newRate - New rate from vendor bill
         */
        function updateItemReceiptLineRate(irId, poLineId, itemId, newRate) {
            log.debug('Updating IR Line', {
                irId: irId,
                itemId: itemId,
                newRate: newRate
            });

            // Load the Item Receipt in standard mode
            var irRecord = record.load({
                type: record.Type.ITEM_RECEIPT,
                id: irId,
                isDynamic: false
            });

            // Find all matching lines by item ID and update them
            var lineCount = irRecord.getLineCount({ sublistId: 'item' });
            var linesUpdated = 0;

            log.debug('IR Line Count', 'Total lines: ' + lineCount);

            for (var i = 0; i < lineCount; i++) {
                var currentItem = irRecord.getSublistValue({
                    sublistId: 'item',
                    fieldId: 'item',
                    line: i
                });

                var currentRate = irRecord.getSublistValue({
                    sublistId: 'item',
                    fieldId: 'rate',
                    line: i
                });

                log.debug('Line ' + i + ' Info', {
                    item: currentItem,
                    rate: currentRate,
                    searchingForItem: itemId
                });

                // Match by item ID - update all instances of this item
                if (currentItem && currentItem.toString() === itemId.toString()) {
                    var oldRate = irRecord.getSublistValue({
                        sublistId: 'item',
                        fieldId: 'rate',
                        line: i
                    });

                    // Update the rate
                    irRecord.setSublistValue({
                        sublistId: 'item',
                        fieldId: 'rate',
                        line: i,
                        value: parseFloat(newRate)
                    });

                    linesUpdated++;
                    log.debug('Line Found and Updated', {
                        line: i,
                        item: currentItem,
                        oldRate: oldRate,
                        newRate: newRate
                    });
                }
            }

            if (linesUpdated === 0) {
                throw new Error('No lines found with Item ID ' + itemId + ' on IR ' + irId);
            }

            // Save the record
            var savedId = irRecord.save({
                enableSourcing: false,
                ignoreMandatoryFields: true
            });

            log.audit('IR Updated Successfully', {
                irId: savedId,
                itemId: itemId,
                linesUpdated: linesUpdated,
                newRate: newRate
            });

            return savedId;
        }

        /**
        * Builds the main page HTML content
        * @param {Object} params - URL parameters
        * @returns {string} HTML content
        */
        function buildPageHTML(params) {
            var html = '<style>' + getStyles() + '</style>';
            html += '<div class="container">';

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

                // Auto-submit to continue processing
                html += '<form id="continueForm" method="POST" style="display:none;">';
                html += '<input type="hidden" name="selected_variances" value="' + escapeHtml(params.selected_variances || '') + '" />';
                html += '<input type="hidden" name="batch_index" value="' + escapeHtml(params.batch_index || '0') + '" />';
                html += '<input type="hidden" name="success_count" value="' + escapeHtml(params.success_count || '0') + '" />';
                html += '<input type="hidden" name="error_count" value="' + escapeHtml(params.error_count || '0') + '" />';
                html += '<input type="hidden" name="previous_errors" value="' + escapeHtml(params.previous_errors || '[]') + '" />';
                html += '<input type="hidden" name="previous_updated" value="' + escapeHtml(params.previous_updated || '[]') + '" />';
                html += '</form>';
                html += '<script>setTimeout(function() { document.getElementById("continueForm").submit(); }, 1000);</script>';
                html += '</div>';
                return html;
            }

            // Show success message if redirected after update
            if (params.updateSuccess === 'true') {
                html += buildUpdateSuccessMessage(params);
            }

            if (params.adjustmentSuccess === 'true') {
                html += '<div class="success-message">';
                html += '<strong>✓ Closed Period Adjustment Complete</strong><br />';
                html += 'Vendor Bill <strong>' + escapeHtml(params.vbNumber) + '</strong> updated<br />';
                html += 'Item: ' + escapeHtml(params.itemName) + '<br />';
                html += 'Adjustment Amount: $' + params.adjustmentAmount + '<br />';
                html += 'Journal Entry <strong>' + escapeHtml(params.jeNumber) + '</strong> created';
                html += '</div>';
            }

            // Show error message if present
            if (params.error) {
                html += '<div class="error-message">';
                html += '<strong>✗ Error</strong><br />';
                html += escapeHtml(params.error);
                html += '</div>';
            }

            // Add instructions
            html += '<div class="instructions">';
            html += '<h3>Instructions</h3>';
            html += '<p>This tool identifies Item Receipts where the line rate differs from the corresponding Vendor Bill rate by $0.01 or more.</p>';
            html += '<p><strong>Process:</strong></p>';
            html += '<ul>';
            html += '<li>Review variances below (sorted by oldest PO date first)</li>';
            html += '<li>Select the Item Receipt lines you want to update</li>';
            html += '<li>Click "Update Selected Item Receipts" to change IR rates to match VB rates</li>';
            html += '<li>The Accrued Purchases account will be adjusted automatically</li>';
            html += '</ul>';
            html += '<p><strong>Note:</strong> When multiple Vendor Bills exist for the same PO line, the oldest VB is matched to the oldest IR to ensure balanced accounting.</p>';
            html += '</div>';

            // Get variance data
            var variancePairs = getVariancePairs();

            if (variancePairs.length === 0) {
                html += '<div class="info-message">';
                html += '<strong>ℹ No Variances Found</strong><br />';
                html += 'All Item Receipt rates match their corresponding Vendor Bill rates.';
                html += '</div>';
            } else {
                html += '<div class="summary-info">';
                html += '<strong>Total Variances Found:</strong> ' + variancePairs.length + ' line(s)';
                html += '</div>';
                html += buildVarianceTable(variancePairs);
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
            html += '<strong>✓ Update Process Complete</strong><br />';
            html += params.successCount + ' line(s) updated successfully.';

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
                        html += '<th style="text-align: left; padding: 8px;">IR #</th>';
                        html += '<th style="text-align: left; padding: 8px;">Item</th>';
                        html += '<th style="text-align: left; padding: 8px;">Reason</th>';
                        html += '</tr>';
                        html += '</thead>';
                        html += '<tbody>';
                        errors.forEach(function (err) {
                            html += '<tr style="border-bottom: 1px solid #f5c6cb;">';
                            html += '<td style="padding: 8px;"><a href="/app/accounting/transactions/itemrcpt.nl?id=' + err.irId + '" target="_blank">' + escapeHtml(err.irNumber) + '</a></td>';
                            html += '<td style="padding: 8px;">' + escapeHtml(err.itemName) + '</td>';
                            html += '<td style="padding: 8px; color: #721c24;">' + escapeHtml(err.error) + '</td>';
                            html += '</tr>';
                        });
                        html += '</tbody>';
                        html += '</table>';
                        html += '</div>';

                        // Add summary of common errors
                        var closedPeriodErrors = errors.filter(function (e) { return e.error.indexOf('closed') !== -1; }).length;
                        var notFoundErrors = errors.filter(function (e) { return e.error.indexOf('not found') !== -1; }).length;

                        if (closedPeriodErrors > 0 || notFoundErrors > 0) {
                            html += '<div style="margin-top: 10px; padding: 10px; background: #fff3cd; border: 1px solid #ffc107; border-radius: 4px; font-size: 13px;">';
                            html += '<strong>Common Issues:</strong><ul style="margin: 5px 0; padding-left: 20px;">';
                            if (closedPeriodErrors > 0) {
                                html += '<li><strong>' + closedPeriodErrors + ' record(s)</strong> in closed accounting periods - please reopen period or contact accounting</li>';
                            }
                            if (notFoundErrors > 0) {
                                html += '<li><strong>' + notFoundErrors + ' record(s)</strong> where item was not found - data may have changed</li>';
                            }
                            html += '</ul></div>';
                        }
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
                        html += '<strong>✓ Successfully Updated Records:</strong>';
                        html += '<div style="max-height: 200px; overflow-y: auto; margin-top: 5px;">';
                        html += '<ul style="margin: 0; padding-left: 20px;">';
                        updated.forEach(function (rec) {
                            html += '<li>' + escapeHtml(rec.irNumber) + ' - ' + escapeHtml(rec.itemName) + ' → $' + parseFloat(rec.newRate).toFixed(2) + '</li>';
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
   * @returns {string} HTML table content
   */
        function buildVarianceTable(variancePairs) {
            var html = '<form id="varianceForm" method="POST">';
            html += '<table class="variance-table">';
            html += '<thead>';
            html += '<tr>';
            html += '<th><input type="checkbox" id="selectAll" title="Select/Deselect All" /></th>';
            html += '<th>PO #</th>';
            html += '<th>PO Date</th>';
            html += '<th>Vendor</th>';
            html += '<th>Item</th>';
            html += '<th>IR #</th>';
            html += '<th>IR Date</th>';
            html += '<th>Period</th>';
            html += '<th class="rate-cell">IR Rate</th>';
            html += '<th>VB #</th>';
            html += '<th>VB Date</th>';
            html += '<th>VB Rate</th>';
            html += '<th>Variance</th>';
            html += '<th>Actions</th>';
            html += '</tr>';
            html += '</thead>';
            html += '<tbody>';

            variancePairs.forEach(function (pair) {
                var variance = pair.vb_rate - pair.ir_rate;
                var varianceClass = Math.abs(variance) >= 0.01 ? 'has-variance' : '';
                var isPeriodClosed = pair.ir_period_closed;

                var checkboxValue = pair.ir_id + '|' +
                    pair.ir_line_id + '|' +
                    pair.vb_rate.toFixed(2) + '|' +
                    pair.ir_number + '|' +
                    pair.item_name + '|' +
                    pair.item_id;

                html += '<tr' + (isPeriodClosed ? ' class="closed-period-row"' : '') + '>';

                // Checkbox - disabled if period is closed
                html += '<td><input type="checkbox" class="variance-checkbox" value="' + escapeHtml(checkboxValue) + '"' +
                    (isPeriodClosed ? ' disabled title="Period is closed"' : '') + ' /></td>';

                html += '<td><a href="/app/accounting/transactions/purchord.nl?id=' + pair.po_id + '" target="_blank">' + escapeHtml(pair.po_number) + '</a></td>';
                html += '<td>' + formatDate(pair.po_date) + '</td>';
                html += '<td>' + escapeHtml(pair.vendor_name) + '</td>';
                html += '<td>' + escapeHtml(pair.item_name) + '</td>';
                html += '<td><a href="/app/accounting/transactions/itemrcpt.nl?id=' + pair.ir_id + '" target="_blank">' + escapeHtml(pair.ir_number) + '</a></td>';
                html += '<td>' + formatDate(pair.ir_date) + '</td>';
                html += '<td><span class="period-status ' + (isPeriodClosed ? 'period-closed' : 'period-open') + '">' +
                    (isPeriodClosed ? '🔒 Closed' : '✓ Open') + '</span></td>';
                html += '<td class="rate-cell">$' + pair.ir_rate.toFixed(2) + '</td>';
                html += '<td><a href="/app/accounting/transactions/vendbill.nl?id=' + pair.vb_id + '" target="_blank">' + escapeHtml(pair.vb_number) + '</a></td>';
                html += '<td>' + formatDate(pair.vb_date) + '</td>';
                html += '<td class="rate-cell vb-rate">$' + pair.vb_rate.toFixed(2) + '</td>';
                html += '<td class="variance-cell ' + varianceClass + '">$' + variance.toFixed(2) + '</td>';

                // Action button - only enabled if period is closed
                html += '<td><button type="button" class="action-button' + (isPeriodClosed ? '' : ' action-button-disabled') + '"' +
                    (isPeriodClosed ? '' : ' disabled') +
                    ' onclick="handleClosedPeriodAdjustment(\'' +
                    escapeHtml(pair.vb_id) + '\',\'' +
                    escapeHtml(pair.item_id) + '\',\'' +
                    pair.vb_rate.toFixed(2) + '\',\'' +
                    pair.ir_rate.toFixed(2) + '\',\'' +
                    escapeHtml(pair.vb_number) + '\',\'' +
                    escapeHtml(pair.item_name) + '\')">' +
                    'Closed Period Adjustment</button></td>';

                html += '</tr>';
            });

            html += '</tbody>';
            html += '</table>';

            html += '<div class="button-container">';
            html += '<button type="button" class="submit-button" onclick="submitVariances()">Update Selected Item Receipts</button>';
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
         * Gets variance pairs by querying and matching oldest IR to oldest VB
         * @returns {Array} Array of variance pair objects
         */
        function getVariancePairs() {
            var rawResults = searchIRVBVariances();
            var poLineGroups = groupByPOLine(rawResults);
            var variancePairs = createVariancePairs(poLineGroups);

            log.debug('Variance Pairs Created', 'Total pairs: ' + variancePairs.length);

            return variancePairs;
        }

        /**
  * Searches for IR/VB rate variances using saved search
  * @returns {Array} Raw search results
  */
        function searchIRVBVariances() {
            var varianceSearch = search.create({
                type: search.Type.TRANSACTION,
                filters: [
                    ['type', 'anyof', 'PurchOrd'],
                    'AND',
                    ['mainline', 'is', 'F'],
                    'AND',
                    ['billingtransaction.quantity', 'greaterthan', '0'],
                    'AND',
                    ['fulfillingtransaction.quantity', 'greaterthan', '0'],
                    'AND',
                    ['formulanumeric: NVL({fulfillingtransaction.rate},0)-NVL({billingtransaction.amount}/{billingtransaction.quantity},0)', 'notequalto', '0']
                ],
                columns: [
                    search.createColumn({ name: 'internalid', label: 'PO ID' }),
                    search.createColumn({ name: 'tranid', label: 'PO Number' }),
                    search.createColumn({ name: 'trandate', label: 'PO Date', sort: search.Sort.ASC }),
                    search.createColumn({ name: 'entity', label: 'Vendor ID' }),
                    search.createColumn({ name: 'entityid', join: 'vendor', label: 'Vendor Name' }),
                    search.createColumn({ name: 'altname', join: 'vendor', label: 'Vendor Display Name' }),
                    search.createColumn({ name: 'lineuniquekey', label: 'PO Line ID' }),
                    search.createColumn({ name: 'item', label: 'Item ID' }),
                    search.createColumn({ name: 'itemid', join: 'item', label: 'Item Number' }),
                    search.createColumn({ name: 'displayname', join: 'item', label: 'Item Name' }),
                    search.createColumn({ name: 'rate', label: 'PO Rate' }),
                    // Item Receipt columns
                    search.createColumn({ name: 'internalid', join: 'fulfillingtransaction', label: 'IR ID' }),
                    search.createColumn({ name: 'tranid', join: 'fulfillingtransaction', label: 'IR Number' }),
                    search.createColumn({ name: 'trandate', join: 'fulfillingtransaction', label: 'IR Date' }),
                    search.createColumn({ name: 'postingperiod', join: 'fulfillingtransaction', label: 'IR Period' }),
                    search.createColumn({ name: 'lineuniquekey', join: 'fulfillingtransaction', label: 'IR Line ID' }),
                    search.createColumn({ name: 'quantity', join: 'fulfillingtransaction', label: 'IR Quantity' }),
                    search.createColumn({ name: 'rate', join: 'fulfillingtransaction', label: 'IR Rate' }),
                    // Vendor Bill columns
                    search.createColumn({ name: 'internalid', join: 'billingtransaction', label: 'VB ID' }),
                    search.createColumn({ name: 'tranid', join: 'billingtransaction', label: 'VB Number' }),
                    search.createColumn({ name: 'trandate', join: 'billingtransaction', label: 'VB Date' }),
                    search.createColumn({ name: 'lineuniquekey', join: 'billingtransaction', label: 'VB Line ID' }),
                    search.createColumn({ name: 'quantity', join: 'billingtransaction', label: 'VB Quantity' }),
                    search.createColumn({ name: 'rate', join: 'billingtransaction', label: 'VB Rate' })
                ]
            });

            var results = [];
            varianceSearch.run().each(function (result) {
                var itemName = result.getText({ name: 'item' }) || result.getValue({ name: 'displayname', join: 'item' }) || '';
                var itemNumber = result.getValue({ name: 'item' }) || '';
                var vendorName = result.getValue({ name: 'altname', join: 'vendor' }) ||
                    result.getValue({ name: 'entityid', join: 'vendor' }) ||
                    result.getText({ name: 'entity' }) ||
                    'Unknown Vendor';

                var periodId = result.getValue({ name: 'postingperiod', join: 'fulfillingtransaction' });
                var isPeriodClosed = false;

                // Lookup period status if we have a period ID
                if (periodId) {
                    try {
                        var periodLookup = search.lookupFields({
                            type: search.Type.ACCOUNTING_PERIOD,
                            id: periodId,
                            columns: ['closed', 'alllocked']
                        });
                        isPeriodClosed = periodLookup.closed || periodLookup.alllocked;
                    } catch (e) {
                        log.error('Period Lookup Error', 'Period ID: ' + periodId + ', Error: ' + e.message);
                    }
                }

                results.push({
                    po_id: result.getValue({ name: 'internalid' }),
                    po_number: result.getValue({ name: 'tranid' }),
                    po_date: result.getValue({ name: 'trandate' }),
                    vendor_name: vendorName,
                    po_line_id: result.getValue({ name: 'lineuniquekey' }),
                    item_id: result.getValue({ name: 'item' }),
                    item_number: itemNumber,
                    item_name: itemName,
                    po_rate: result.getValue({ name: 'rate' }),
                    ir_id: result.getValue({ name: 'internalid', join: 'fulfillingtransaction' }),
                    ir_number: result.getValue({ name: 'tranid', join: 'fulfillingtransaction' }),
                    ir_date: result.getValue({ name: 'trandate', join: 'fulfillingtransaction' }),
                    ir_period_id: periodId,
                    ir_period_closed: isPeriodClosed,
                    ir_line_id: result.getValue({ name: 'lineuniquekey', join: 'fulfillingtransaction' }),
                    ir_quantity: result.getValue({ name: 'quantity', join: 'fulfillingtransaction' }),
                    ir_rate: result.getValue({ name: 'rate', join: 'fulfillingtransaction' }),
                    vb_id: result.getValue({ name: 'internalid', join: 'billingtransaction' }),
                    vb_number: result.getValue({ name: 'tranid', join: 'billingtransaction' }),
                    vb_date: result.getValue({ name: 'trandate', join: 'billingtransaction' }),
                    vb_line_id: result.getValue({ name: 'lineuniquekey', join: 'billingtransaction' }),
                    vb_quantity: result.getValue({ name: 'quantity', join: 'billingtransaction' }),
                    vb_rate: result.getValue({ name: 'rate', join: 'billingtransaction' })
                });
                return true;
            });

            log.debug('Search Results', 'Total rows: ' + results.length);
            return results;
        }

        /**
         * Groups raw query results by PO Line ID
         * @param {Array} rawResults - Raw search results
         * @returns {Object} Grouped results by PO line
         */
        function groupByPOLine(rawResults) {
            var groups = {};

            rawResults.forEach(function (row) {
                var poLineKey = row.po_line_id;

                if (!groups[poLineKey]) {
                    groups[poLineKey] = {
                        poInfo: {
                            po_id: row.po_id,
                            po_number: row.po_number,
                            po_date: row.po_date,
                            vendor_name: row.vendor_name,
                            item_id: row.item_id,
                            item_number: row.item_number,
                            item_name: row.item_name
                        },
                        itemReceipts: [],
                        vendorBills: []
                    };
                }

                var irExists = groups[poLineKey].itemReceipts.some(function (ir) {
                    return ir.ir_line_id === row.ir_line_id;
                });
                if (!irExists) {
                    groups[poLineKey].itemReceipts.push({
                        ir_id: row.ir_id,
                        ir_number: row.ir_number,
                        ir_date: row.ir_date,
                        ir_period_closed: row.ir_period_closed,
                        ir_line_id: row.ir_line_id,
                        ir_quantity: parseFloat(row.ir_quantity),
                        ir_rate: parseFloat(row.ir_rate)
                    });
                }

                var vbExists = groups[poLineKey].vendorBills.some(function (vb) {
                    return vb.vb_line_id === row.vb_line_id;
                });
                if (!vbExists) {
                    groups[poLineKey].vendorBills.push({
                        vb_id: row.vb_id,
                        vb_number: row.vb_number,
                        vb_date: row.vb_date,
                        vb_line_id: row.vb_line_id,
                        vb_quantity: parseFloat(row.vb_quantity),
                        vb_rate: parseFloat(row.vb_rate)
                    });
                }
            });

            log.debug('Grouped by PO Line', 'Total PO lines with variances: ' + Object.keys(groups).length);
            return groups;
        }

        /**
         * Creates variance pairs by matching oldest IR to oldest VB
         * @param {Object} poLineGroups - Grouped results
         * @returns {Array} Array of variance pair objects
         */
        function createVariancePairs(poLineGroups) {
            var pairs = [];

            Object.keys(poLineGroups).forEach(function (poLineKey) {
                var group = poLineGroups[poLineKey];

                group.itemReceipts.sort(function (a, b) {
                    return new Date(a.ir_date) - new Date(b.ir_date);
                });
                group.vendorBills.sort(function (a, b) {
                    return new Date(a.vb_date) - new Date(b.vb_date);
                });

                var maxPairs = Math.max(group.itemReceipts.length, group.vendorBills.length);

                for (var i = 0; i < maxPairs; i++) {
                    var ir = group.itemReceipts[i];
                    var vb = group.vendorBills[i];

                    if (ir && vb) {
                        var variance = vb.vb_rate - ir.ir_rate;

                        if (Math.abs(variance) >= 0.01) {
                            pairs.push({
                                po_id: group.poInfo.po_id,
                                po_number: group.poInfo.po_number,
                                po_date: group.poInfo.po_date,
                                vendor_name: group.poInfo.vendor_name,
                                item_id: group.poInfo.item_id,
                                item_number: group.poInfo.item_number,
                                item_name: group.poInfo.item_name,
                                ir_id: ir.ir_id,
                                ir_number: ir.ir_number,
                                ir_date: ir.ir_date,
                                ir_period_closed: ir.ir_period_closed,
                                ir_line_id: ir.ir_line_id,
                                ir_quantity: ir.ir_quantity,
                                ir_rate: ir.ir_rate,
                                vb_id: vb.vb_id,
                                vb_number: vb.vb_number,
                                vb_date: vb.vb_date,
                                vb_line_id: vb.vb_line_id,
                                vb_quantity: vb.vb_quantity,
                                vb_rate: vb.vb_rate
                            });
                        }
                    }
                }
            });

            return pairs;
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
        
        .instructions {
            background: white;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .instructions h3 {
            margin-top: 0;
            color: #1a73e8;
        }
        
        .instructions ul {
            margin: 10px 0;
            padding-left: 20px;
        }
        
        .instructions li {
            margin: 5px 0;
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
        
        .closed-period-row {
            background: #fff3cd !important;
        }
        
        .closed-period-row:hover {
            background: #ffe69c !important;
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
        
        .rate-cell {
            text-align: right;
            font-family: Arial, sans-serif;
            font-size: 14px;
        }
        
        .vb-rate {
            font-weight: bold;
            color: #2e7d32;
        }
        
        .variance-cell {
            text-align: right;
            font-family: Arial, sans-serif;
            font-weight: bold;
            font-size: 14px;
        }
        
        .has-variance {
            color: #d32f2f;
        }
        
        .period-status {
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: 600;
            white-space: nowrap;
        }
        
        .period-open {
            background: #d4edda;
            color: #155724;
        }
        
        .period-closed {
            background: #f8d7da;
            color: #721c24;
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
        
        .variance-checkbox:disabled {
            cursor: not-allowed;
            opacity: 0.5;
        }

        .action-button {
            background: #f57c00;
            color: white;
            border: none;
            padding: 6px 12px;
            font-size: 12px;
            font-weight: 600;
            border-radius: 4px;
            cursor: pointer;
            transition: background 0.2s;
            white-space: nowrap;
        }

        .action-button:hover:not(:disabled) {
            background: #e65100;
        }

        .action-button:active:not(:disabled) {
            background: #d84315;
        }
        
        .action-button-disabled {
            background: #ccc !important;
            cursor: not-allowed !important;
            opacity: 0.5;
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
                    var checkboxes = document.querySelectorAll('.variance-checkbox:not(:disabled)');
                    checkboxes.forEach(function(cb) {
                        cb.checked = selectAll.checked;
                    });
                });
            }
        });
        
        function submitVariances() {
            var checkboxes = document.querySelectorAll('.variance-checkbox:checked:not(:disabled)');
            
            if (checkboxes.length === 0) {
                alert('Please select at least one variance to update.\\n\\nNote: Item Receipts in closed periods cannot be updated directly. Use the "Closed Period Adjustment" button instead.');
                return;
            }
            
            var confirmMessage = 'Update ' + checkboxes.length + ' Item Receipt line(s) with Vendor Bill rates?\\n\\n';
            confirmMessage += 'This will change the Item Receipt rate to match the Vendor Bill rate.\\n';
            confirmMessage += 'This action cannot be undone.\\n\\n';
            confirmMessage += 'Continue?';
            
            if (!confirm(confirmMessage)) {
                return;
            }
            
            var selected = [];
            checkboxes.forEach(function(cb) {
                selected.push(cb.value);
            });
            
            var form = document.getElementById('varianceForm');
            if (!form) {
                form = document.querySelector('form[id="varianceForm"]');
            }
            if (!form) {
                form = document.querySelector('form');
            }
            
            if (!form) {
                alert('Error: Could not find form element. Please refresh and try again.');
                return;
            }
            
            var input = document.createElement('input');
            input.type = 'hidden';
            input.name = 'selected_variances';
            input.value = selected.join(',');
            form.appendChild(input);
            
            var submitButton = document.querySelector('.submit-button');
            if (submitButton) {
                submitButton.disabled = true;
                submitButton.textContent = 'Updating ' + checkboxes.length + ' record(s)...';
            }
            
            form.submit();
        }
        
        function handleClosedPeriodAdjustment(vbId, itemId, vbRate, irRate, vbNumber, itemName) {
            var confirmMsg = 'Perform Closed Period Adjustment?\\n\\n';
            confirmMsg += 'VB: ' + vbNumber + '\\n';
            confirmMsg += 'Item: ' + itemName + '\\n';
            confirmMsg += 'Current VB Rate: $' + parseFloat(vbRate).toFixed(2) + '\\n';
            confirmMsg += 'IR Rate (target): $' + parseFloat(irRate).toFixed(2) + '\\n';
            confirmMsg += 'Adjustment: $' + (parseFloat(vbRate) - parseFloat(irRate)).toFixed(2) + '\\n\\n';
            confirmMsg += 'This will:\\n';
            confirmMsg += '1. Update VB item rate to match IR\\n';
            confirmMsg += '2. Add expense line to Accrued Purchases\\n';
            confirmMsg += '3. Create offsetting JE\\n\\n';
            confirmMsg += 'Continue?';
            
            if (!confirm(confirmMsg)) {
                return;
            }
            
            var form = document.createElement('form');
            form.method = 'POST';
            form.style.display = 'none';
            
            var inputs = {
                action: 'closed_period_adjustment',
                vb_id: vbId,
                item_id: itemId,
                vb_rate: vbRate,
                ir_rate: irRate,
                vb_number: vbNumber,
                item_name: itemName
            };
            
            for (var key in inputs) {
                var input = document.createElement('input');
                input.type = 'hidden';
                input.name = key;
                input.value = inputs[key];
                form.appendChild(input);
            }
            
            document.body.appendChild(form);
            form.submit();
        }
    `;
        }

        return {
            onRequest: onRequest
        };
    });