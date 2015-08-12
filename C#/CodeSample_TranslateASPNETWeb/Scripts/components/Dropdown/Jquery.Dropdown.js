/**
 * Dropdown Plugin
 * 
 * Given .ms-Dropdown containers with generic <select> elements inside, this plugin hides the original
 * dropdown and creates a new "fake" dropdown that can more easily be styled across browsers.
 * 
 * @param  {jQuery Object}  One or more .ms-Dropdown containers, each with a dropdown (.ms-Dropdown-select)
 * @return {jQuery Object}  The same containers (allows for chaining)
 */
(function ($) {
    $.fn.Dropdown = function () {

        /** Go through each dropdown we've been given. */
        return this.each(function () {

            var $dropdownWrapper = $(this),
                $originalDropdown = $dropdownWrapper.children('.ms-Dropdown-select'),
                $originalDropdownOptions = $originalDropdown.children('option'),
                originalDropdownID = this.id,
                newDropdownTitle = '',
                newDropdownItems = '',
                newDropdownSource = '';

            /** Go through the options to fill up newDropdownTitle and newDropdownItems. */
            $originalDropdownOptions.each(function (index, option) {
        
                /** If the option is selected, it should be the new dropdown's title. */
                if (option.selected) {
                    newDropdownTitle = option.text;
                }

                /** Add this option to the list of items. */
                newDropdownItems += '<li class="ms-Dropdown-item' + ( (option.disabled) ? ' is-disabled"' : '"' ) + '>' + option.text + '</li>';
            
            });

            /** Insert the replacement dropdown. */
            newDropdownSource = '<span class="ms-Dropdown-title">' + newDropdownTitle + '</span><ul class="ms-Dropdown-items">' + newDropdownItems + '</ul>';
            $dropdownWrapper.append(newDropdownSource);

            function _openDropdown() {
                if (!$dropdownWrapper.hasClass('is-disabled')) {

                    /** First, let's close any open dropdowns on this page. */
                    $dropdownWrapper.find('.ms-Dropdown--open').removeClass('ms-Dropdown--open');

                    /** Stop the click event from propagating, which would just close the dropdown immediately. */
                    event.stopPropagation();

                    /** Before opening, size the items list to match the dropdown. */
                    var dropdownWidth = $(this).parents(".ms-Dropdown").width();
                    $(this).next(".ms-Dropdown-items").css('width', dropdownWidth + 'px');
                
                    /** Go ahead and open that dropdown. */
                    $dropdownWrapper.toggleClass('ms-Dropdown--open');

                    /** Temporarily bind an event to the document that will close this dropdown when clicking anywhere. */
                    $(document).bind("click.dropdown", function(event) {
                        $dropdownWrapper.removeClass('ms-Dropdown--open');
                        $(document).unbind('click.dropdown');
                    });
                }
            };

            /** Toggle open/closed state of the dropdown when clicking its title. */
            $dropdownWrapper.on('click', '.ms-Dropdown-title', function(event) {
                _openDropdown();
            });

            /** Keyboard accessibility */
            $dropdownWrapper.on('keyup', function(e) {
                var keyCode = e.keyCode || e.which;
                // Open dropdown on enter or arrow up or arrow down and focus on first option
                if (!$(this).hasClass('ms-Dropdown--open')) {
                    if (keyCode === 13 || keyCode === 38 || keyCode === 40) {
                       _openDropdown();
                       if (!$(this).find('.ms-Dropdown-item').hasClass('ms-Dropdown-item--selected')) {
                        $(this).find('.ms-Dropdown-item:first').addClass('ms-Dropdown-item--selected');
                       }
                    }
                }
                else if ($(this).hasClass('ms-Dropdown--open')) {
                    // Up arrow focuses previous option
                    if (keyCode === 38) {
                        if ($(this).find('.ms-Dropdown-item.ms-Dropdown-item--selected').prev().siblings().size() > 0) {
                            $(this).find('.ms-Dropdown-item.ms-Dropdown-item--selected').removeClass('ms-Dropdown-item--selected').prev().addClass('ms-Dropdown-item--selected');
                        }
                    }
                    // Down arrow focuses next option
                    if (keyCode === 40) {
                        if ($(this).find('.ms-Dropdown-item.ms-Dropdown-item--selected').next().siblings().size() > 0) {
                            $(this).find('.ms-Dropdown-item.ms-Dropdown-item--selected').removeClass('ms-Dropdown-item--selected').next().addClass('ms-Dropdown-item--selected');
                        }
                    }
                    // Enter to select item
                    if (keyCode === 13) {
                        if (!$dropdownWrapper.hasClass('is-disabled')) {

                            // Item text
                            var selectedItemText = $(this).find('.ms-Dropdown-item.ms-Dropdown-item--selected').text()

                            $(this).find('.ms-Dropdown-title').html(selectedItemText);

                            /** Update the original dropdown. */
                            $originalDropdown.find("option").each(function(key, value) {
                                if (value.text === selectedItemText) {
                                    $(this).prop('selected', true);
                                } else {
                                    $(this).prop('selected', false);
                                }
                            });
                            $originalDropdown.change();

                            $(this).removeClass('ms-Dropdown--open');
                        }
                    }                
                }

                // Close dropdown on esc
                if (keyCode === 27) {
                    $(this).removeClass('ms-Dropdown--open');
                }
            });

            /** Select an option from the dropdown. */
            $dropdownWrapper.on('click', '.ms-Dropdown-item', function () {
                if (!$dropdownWrapper.hasClass('is-disabled')) {

                    /** Deselect all items and select this one. */
                    $(this).siblings('.ms-Dropdown-item').removeClass('ms-Dropdown-item--selected')
                    $(this).addClass('ms-Dropdown-item--selected');

                    /** Update the replacement dropdown's title. */
                    $(this).parents().siblings('.ms-Dropdown-title').html($(this).text());

                    /** Update the original dropdown. */
                    var selectedItemText = $(this).text();
                    $originalDropdown.find("option").each(function(key, value) {
                        if (value.text === selectedItemText) {
                            $(this).prop('selected', true);
                        } else {
                            $(this).prop('selected', false);
                        }
                    });
                    $originalDropdown.change();
                }
            });

        });
    };
})(jQuery);