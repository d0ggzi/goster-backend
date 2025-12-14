from fastapi import Response, Request


async def catch_exceptions_middleware(request: Request, call_next):
    try:
        return await call_next(request)
    except Exception as e:
        print(f"Exception raised: {e}")
        return Response("Internal server error", status_code=500)
