from fastapi import Header


async def get_current_user(x_user_id: str = Header(None)) -> str:
    return x_user_id
