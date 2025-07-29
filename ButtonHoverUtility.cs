using System;
using System.Windows.Forms;

namespace my_addin
{
    /// <summary>
    /// Utility class for adding consistent hover effects to buttons
    /// Eliminates the need for hundreds of duplicate MouseEnter/MouseLeave event handlers
    /// </summary>
    public static class ButtonHoverUtility
    {
        /// <summary>
        /// Adds standard hover effect to a button (border appears on hover)
        /// </summary>
        /// <param name="button">The button to add hover effect to</param>
        /// <param name="hoverBorderSize">Border size when hovering (default: 1)</param>
        /// <param name="normalBorderSize">Border size when not hovering (default: 0)</param>
        public static void EnableHoverEffect(this Button button, int hoverBorderSize = 1, int normalBorderSize = 0)
        {
            if (button == null) return;

            // Remove any existing hover handlers to prevent duplicates
            button.MouseEnter -= Button_MouseEnter;
            button.MouseLeave -= Button_MouseLeave;

            // Store the border sizes in the button's Tag property
            button.Tag = new HoverSettings { HoverBorderSize = hoverBorderSize, NormalBorderSize = normalBorderSize };

            // Add the generic event handlers
            button.MouseEnter += Button_MouseEnter;
            button.MouseLeave += Button_MouseLeave;
        }

        /// <summary>
        /// Removes hover effect from a button
        /// </summary>
        /// <param name="button">The button to remove hover effect from</param>
        public static void DisableHoverEffect(this Button button)
        {
            if (button == null) return;

            button.MouseEnter -= Button_MouseEnter;
            button.MouseLeave -= Button_MouseLeave;
            button.Tag = null;
        }

        /// <summary>
        /// Applies hover effects to multiple buttons at once
        /// </summary>
        /// <param name="buttons">Array of buttons to apply hover effects to</param>
        public static void EnableHoverEffects(params Button[] buttons)
        {
            foreach (var button in buttons)
            {
                button?.EnableHoverEffect();
            }
        }

        /// <summary>
        /// Applies hover effects to all buttons in a container control
        /// </summary>
        /// <param name="container">Container control to search for buttons</param>
        /// <param name="recursive">Whether to search recursively in child containers</param>
        public static void EnableHoverEffectsForContainer(Control container, bool recursive = true)
        {
            if (container == null) return;

            foreach (Control control in container.Controls)
            {
                if (control is Button button)
                {
                    button.EnableHoverEffect();
                }
                else if (recursive && control.HasChildren)
                {
                    EnableHoverEffectsForContainer(control, recursive);
                }
            }
        }

        #region Private Implementation

        private static void Button_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button button && button.Tag is HoverSettings settings)
            {
                button.FlatAppearance.BorderSize = settings.HoverBorderSize;
            }
        }

        private static void Button_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button button && button.Tag is HoverSettings settings)
            {
                button.FlatAppearance.BorderSize = settings.NormalBorderSize;
            }
        }

        private class HoverSettings
        {
            public int HoverBorderSize { get; set; }
            public int NormalBorderSize { get; set; }
        }

        #endregion
    }
}