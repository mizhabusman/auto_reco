"""
prompt.py — All prompts live here.

Minimal by design: Claude is the CA expert, we hand it the two ledgers and ask
it to reconcile them. No schema, no forced format — it decides everything.
"""

# The persona Claude takes on.
SYSTEM_PROMPT = "You are an expert Chartered Accountant."

# The instruction placed before the two ledgers.
TASK_INSTRUCTION = "Reconcile these two ledgers."
