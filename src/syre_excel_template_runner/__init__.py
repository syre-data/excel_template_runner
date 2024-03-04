# SPDX-FileCopyrightText: 2024-present Brian Carlsen <carlsen.bri@gmail.com>
#
# SPDX-License-Identifier: MIT
import sys

if sys.version_info < (3, 9):
    _LEGACY_ = True
else:
    _LEGACY_ = False
