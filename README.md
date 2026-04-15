AutoSwitch

AutoSwitch is a MelonLoader mod for Data Center that automatically groups adjacent supported switches into one logical fabric and creates safe in-game LACP bundles where valid.

Features
Automatically detects adjacent switch stacks
Treats grouped switches as one larger switching fabric
Creates safe native LACP bundles for valid multi-cable exits
Prevents bad cross-fabric managed-switch bundles that can break traffic sharing
Uses a reload wake workaround by power-cycling one random eligible switch (up to 30 seconds after scene load)
No player-specific hardcoded switch IDs for safer networking logic

All switches supproted and can be mixed

Installation
Install MelonLoader for Data Center
Put AutoSwitch.dll into your Mods folder
Launch the game
Build and rack switches normally

How It Works
AutoSwitch scans placed switches
Adjacent supported switches are grouped into a shared fabric
Real external multi-cable exits are checked for safe LACP creation
On reload, the mod triggers a small random switch power cycle to wake the network if needed

Notes
Best results come from clean rack layouts and clean cabling
Very large merged fabrics may cause some lag
Cross-fabric bridge experiments can still create strange behavior
Save backups are recommended before major topology changes
Known Limitation

The reload wake behavior is a workaround based on observed game behavior, not an official game system. This can take up to 30 seconds after scene load.

By
Big Texas Jerky
