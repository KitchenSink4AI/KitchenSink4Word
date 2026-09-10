# Changelog

### 2.1.1
- The server reported the framework's version number instead of its own when a client connected. It now reports the version you installed.
- The server introduces itself by name. A connected agent can tell it is talking to KitchenSink4Word, and `get_server_info` now returns the product name, the package name, the landing page, and the rest of the family.
- Published addresses point at kitchensink4.ai.

### 2.1.0
- Update notice: the server checks PyPI for newer releases at most once every seven days, says so in its own payload, and turns off with KS4W_UPDATE_CHECK=off. It never downloads anything.
- Table reads cost about a quarter of what they did; merge fidelity is unchanged and pinned.
- Documents open roughly twice as fast.
- The window-binding layer is consolidated and hardened against a crash class found in live testing.
- The install screen and info card are rewritten in plain language.
