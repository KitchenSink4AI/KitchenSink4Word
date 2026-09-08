# Changelog

### 2.1.0
- Update notice: the server checks PyPI for newer releases at most once every seven days, says so in its own payload, and turns off with KS4W_UPDATE_CHECK=off. It never downloads anything.
- Table reads cost about a quarter of what they did; merge fidelity is unchanged and pinned.
- Documents open roughly twice as fast.
- The window-binding layer is consolidated and hardened against a crash class found in live testing.
- The install screen and info card are rewritten in plain language.
